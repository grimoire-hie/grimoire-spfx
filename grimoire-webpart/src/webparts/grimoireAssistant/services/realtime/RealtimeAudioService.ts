/**
 * RealtimeAudioService
 * WebRTC-based real-time voice conversation with GPT Realtime API.
 * Uses browser-native APIs only (RTCPeerConnection, getUserMedia).
 *
 * Flow:
 * 1. Fetch ephemeral token from proxy /api/realtime/token
 * 2. Create RTCPeerConnection + get user microphone
 * 3. Create data channel for events
 * 4. SDP offer/answer exchange with Foundry endpoint
 * 5. Session.update with tools + instructions
 * 6. Handle function calls, transcripts, and audio
 */

import type { AssistantPlaybackState, IProxyConfig } from '../../store/useGrimoireStore';
import { logService } from '../logging/LogService';
import { normalizeError } from '../utils/errorUtils';

const REALTIME_TOKEN_TIMEOUT_MS = 15_000;
const REALTIME_SDP_TIMEOUT_MS = 20_000;
const RESPONSE_CREATE_RECENT_WINDOW_MS = 220;
const QUEUED_RESPONSE_WATCHDOG_MS = 2_400;
const RESPONSE_CREATED_WATCHDOG_MS = 3_200;
const OUTPUT_AUDIO_BUFFER_STALL_MS = 18_000;
const BARGE_IN_BUFFERING_GRACE_MS = 650;
const DEFERRED_RESPONSE_TOOLS: ReadonlySet<string> = new Set([
  'research_public_web',
  'search_sharepoint',
  'search_people',
  'search_sites',
  'search_emails',
  'show_compose_form',
  'browse_document_library',
  'show_file_details',
  'show_site_info',
  'show_list_items',
  'connect_mcp_server',
  'call_mcp_tool',
  'use_m365_capability',
  'get_my_profile',
  'get_recent_documents',
  'get_trending_documents',
  'recall_notes',
  'read_file_content',
  // Expression changes are UI-only; reopening the response loop can split
  // a spoken reply into an unvoiced text fragment plus a second audio turn.
  'set_expression'
]);

const VOICE_DUPLICATE_DISPLAY_TOOLS: ReadonlySet<string> = new Set([
  'show_markdown',
  'show_info_card',
  'show_selection_list'
]);

// ─── Types ────────────────────────────────────────────────────

export type RealtimeState = 'idle' | 'connecting' | 'connected' | 'speaking' | 'error';

type ResponseCreateReason = 'user-text' | 'context' | 'function-output' | 'queued' | 'deterministic-ack';

export interface IRealtimeCallbacks {
  onStateChange: (state: RealtimeState) => void;
  onTranscript: (text: string, role: 'user' | 'assistant') => void;
  onError: (error: string) => void;
  onAssistantPlaybackStateChange?: (state: AssistantPlaybackState) => void;
  onRemoteStream?: (stream: MediaStream | undefined) => void;
  /**
   * All function calls are delegated here. The handler must return
   * a string, or a Promise resolving to one, to send back as the
   * function call output.
   */
  onFunctionCall: (callId: string, funcName: string, args: Record<string, unknown>) => string | Promise<string>;
  /** Fired when user speech starts or stops (from server VAD). */
  onSpeechActivity?: (activity: 'started' | 'stopped') => void;
}

interface ITokenResponse {
  clientSecret: string;
  expiresAt: string;
  endpoint: string;
}

// ─── Service ──────────────────────────────────────────────────

export class RealtimeAudioService {
  private pc: RTCPeerConnection | undefined;
  private dc: RTCDataChannel | undefined;
  private mediaStream: MediaStream | undefined;
  private audioElement: HTMLAudioElement | undefined;
  private remoteStream: MediaStream | undefined;
  private callbacks: IRealtimeCallbacks | undefined;
  private state: RealtimeState = 'idle';
  private functionCallBuffers: Map<string, string> = new Map();
  /** Accumulates assistant audio transcript deltas (flushed on response.done) */
  private audioTranscriptBuffer: string = '';
  /** Tracks how many session.update messages we've sent for correlating responses */
  private sessionUpdateCount: number = 0;
  /** True while a model response is currently being generated. */
  private responseInFlight: boolean = false;
  /** True when a follow-up response.create should run after the current response completes. */
  private queuedResponseCreate: boolean = false;
  /** Optional custom payload for the next queued response.create message. */
  private queuedResponseCreatePayload: Record<string, unknown> | undefined;
  /** Timestamp when the current response entered the in-flight state. */
  private responseStartedAt: number | undefined;
  /** True when the current response already produced spoken or visible assistant content. */
  private currentResponseHasUserVisibleOutput: boolean = false;
  /** Ordered tool calls observed in the current response. */
  private currentResponseToolCalls: string[] = [];
  /** True when an expression-only response should continue once the current turn finishes. */
  private pendingExpressionOnlyFollowUp: boolean = false;
  /** Timestamp of the last response.create send (used to collapse bursty triggers). */
  private lastResponseCreateAt: number = 0;
  /** True while server output audio buffer is active. */
  private outputAudioBufferActive: boolean = false;
  /** Display tools already emitted in the current user turn. */
  private displayToolsShownInTurn: Set<string> = new Set();
  /** Timestamp when output audio buffer most recently started. */
  private outputAudioBufferStartedAt: number | undefined;
  /** Watchdog for queued response.create calls that can otherwise stall. */
  private queuedResponseWatchdog: ReturnType<typeof setTimeout> | undefined;
  /** Watchdog for a sent response.create that never receives response.created. */
  private pendingResponseCreatedWatchdog: ReturnType<typeof setTimeout> | undefined;
  /** Deferred work moved off the RTC message hot path. */
  private deferredMessageTaskIds: Array<ReturnType<typeof setTimeout>> = [];
  /** Assistant playback state for the active realtime session. */
  private assistantPlaybackState: AssistantPlaybackState = 'idle';

  /**
   * Start a voice session.
   * @param config - Proxy configuration
   * @param callbacks - Event callbacks
   * @param voice - Voice ID (default: 'alloy')
   * @param instructions - System prompt
   * @param tools - Tool definitions for GPT Realtime
   */
  public async connect(
    config: IProxyConfig,
    callbacks: IRealtimeCallbacks,
    voice: string,
    instructions: string,
    tools: object[]
  ): Promise<void> {
    this.callbacks = callbacks;
    this.setState('connecting');

    try {
      // 1. Get ephemeral token from proxy
      const token = await this.fetchToken(config, voice, instructions);

      // 2. Create peer connection
      this.pc = new RTCPeerConnection();

      // 3. Set up remote audio playback
      this.audioElement = new Audio();
      const playbackElement = this.audioElement as HTMLAudioElement & {
        defaultMuted?: boolean;
        playsInline?: boolean;
      };
      playbackElement.autoplay = true;
      playbackElement.muted = false;
      playbackElement.defaultMuted = false;
      playbackElement.playsInline = true;
      this.pc.ontrack = (event) => {
        this.remoteStream = event.streams[0];
        this.callbacks?.onRemoteStream?.(this.remoteStream);
        if (this.audioElement) {
          this.audioElement.srcObject = event.streams[0];
          this.audioElement.play().catch((error: unknown) => {
            const normalizedError = normalizeError(error, 'Realtime remote audio playback failed');
            logService.warning('voice', normalizedError.message);
          });
        }
      };

      // 4. Get microphone and add track
      this.mediaStream = await navigator.mediaDevices.getUserMedia({ audio: true });
      this.mediaStream.getTracks().forEach((track) => {
        this.pc!.addTrack(track, this.mediaStream!);
      });

      // 5. Create data channel for events
      this.dc = this.pc.createDataChannel('oai-events');
      this.dc.onopen = () => this.onDataChannelOpen(tools);
      this.dc.onmessage = (event) => this.onDataChannelMessage(event);
      this.dc.onerror = () => this.handleChannelError();
      this.dc.onclose = () => this.handleChannelFailure('Realtime data channel closed');

      this.pc.onconnectionstatechange = () => {
        const state = this.pc?.connectionState;
        if (!state) return;
        logService.info('voice', `Peer connection state: ${state}`);
        if (state === 'failed') {
          this.handleChannelFailure('Realtime peer connection failed');
        } else if (state === 'closed' && this.state !== 'idle') {
          this.handleChannelFailure('Realtime peer connection closed');
        }
      };
      this.pc.oniceconnectionstatechange = () => {
        const state = this.pc?.iceConnectionState;
        if (!state) return;
        logService.debug('voice', `ICE connection state: ${state}`);
        if (state === 'failed') {
          this.handleChannelFailure('Realtime ICE connection failed');
        }
      };

      // 6. SDP exchange
      const offer = await this.pc.createOffer();
      await this.pc.setLocalDescription(offer);

      const sdpUrl = `${token.endpoint}/openai/v1/realtime/calls`;

      const sdpResponse = await this.fetchWithTimeout(
        sdpUrl,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token.clientSecret}`,
            'Content-Type': 'application/sdp'
          },
          body: offer.sdp
        },
        REALTIME_SDP_TIMEOUT_MS,
        'SDP exchange'
      );

      if (!sdpResponse.ok) {
        throw new Error(`SDP exchange failed: ${sdpResponse.status}`);
      }

      const answerSdp = await sdpResponse.text();
      await this.pc.setRemoteDescription({
        type: 'answer',
        sdp: answerSdp
      });

      this.setState('connected');
      this.setAssistantPlaybackState('idle');
    } catch (error) {
      const normalizedError = normalizeError(error, 'Connection failed');
      this.callbacks?.onError(normalizedError.message);
      this.cleanup();
      this.setAssistantPlaybackState('error');
      this.setState('error');
    }
  }

  /**
   * Send a text message to the ongoing session via data channel.
   * This creates a conversation item with user role text content.
   */
  public sendText(text: string): void {
    if (!this.dc || this.dc.readyState !== 'open') return;
    this.resetVoiceTurnDisplayState();

    // Create a user message in the conversation
    this.dc.send(JSON.stringify({
      type: 'conversation.item.create',
      item: {
        type: 'message',
        role: 'user',
        content: [
          {
            type: 'input_text',
            text
          }
        ]
      }
    }));

    // Trigger the next response from the active realtime session.
    this.requestResponseCreate('user-text');
  }

  /**
   * Send a context message to the LLM conversation.
   * Unlike sendText(), this can optionally skip triggering a response,
   * allowing silent context injection (the LLM absorbs it without speaking).
   */
  public sendContextMessage(text: string, triggerResponse: boolean = false): void {
    if (!this.dc || this.dc.readyState !== 'open') return;

    this.dc.send(JSON.stringify({
      type: 'conversation.item.create',
      item: {
        type: 'message',
        role: 'user',
        content: [{ type: 'input_text', text }]
      }
    }));

    if (triggerResponse) {
      this.requestResponseCreate('context');
    }
  }

  /**
   * Speak a short deterministic acknowledgment in the assistant's realtime voice
   * without adding it to the default conversation history.
   */
  public speakDeterministicText(text: string): boolean {
    if (!this.dc || this.dc.readyState !== 'open') return false;

    const trimmed = text.trim();
    if (!trimmed) return false;

    this.requestResponseCreate('deterministic-ack', {
      conversation: 'none',
      input: [],
      modalities: ['text', 'audio'],
      tool_choice: 'none',
      max_response_output_tokens: 80,
      instructions: `Say exactly this sentence and nothing else: ${JSON.stringify(trimmed)}`
    });

    return true;
  }

  /**
   * Disconnect and clean up all resources.
   */
  public disconnect(): void {
    this.cleanup();
    this.setState('idle');
  }

  /**
   * Mute/unmute the microphone.
   */
  public setMuted(muted: boolean): void {
    if (this.mediaStream) {
      this.mediaStream.getAudioTracks().forEach((track) => {
        track.enabled = !muted;
      });
    }
  }

  public getState(): RealtimeState {
    return this.state;
  }

  /**
   * Get the remote audio stream for lip sync analysis.
   */
  public getRemoteStream(): MediaStream | undefined {
    return this.remoteStream;
  }

  /**
   * Get the local microphone stream for ambient sound analysis.
   */
  public getMicStream(): MediaStream | undefined {
    return this.mediaStream;
  }

  /**
   * Check if the data channel is open and ready.
   */
  public isConnected(): boolean {
    return this.dc?.readyState === 'open';
  }

  public interruptAssistantPlayback(): void {
    if (!this.dc || this.dc.readyState !== 'open') return;

    const bufferingAgeMs = this.responseStartedAt !== undefined
      ? (Date.now() - this.responseStartedAt)
      : Number.POSITIVE_INFINITY;
    const shouldCancelBufferedResponse = this.responseInFlight
      && !this.outputAudioBufferActive
      && bufferingAgeMs >= BARGE_IN_BUFFERING_GRACE_MS;

    if (!shouldCancelBufferedResponse && !this.outputAudioBufferActive) return;

    if (this.responseInFlight) {
      this.dc.send(JSON.stringify({ type: 'response.cancel' }));
    }
    if (this.outputAudioBufferActive) {
      this.dc.send(JSON.stringify({ type: 'output_audio_buffer.clear' }));
    }
    logService.debug('voice', 'Requested realtime assistant playback interruption');
  }

  // ─── Private ──────────────────────────────────────────────────

  private async fetchToken(
    config: IProxyConfig,
    voice: string,
    instructions: string
  ): Promise<ITokenResponse> {
    const response = await this.fetchWithTimeout(
      `${config.proxyUrl}/realtime/token`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-functions-key': config.proxyApiKey
        },
        body: JSON.stringify({ voice, instructions })
      },
      REALTIME_TOKEN_TIMEOUT_MS,
      'realtime token request'
    );

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Token request failed (${response.status}): ${body.slice(0, 200)}`);
    }

    return response.json();
  }

  private onDataChannelOpen(tools: object[]): void {
    // Azure Foundry rejects unknown parameters in session.update and drops
    // the ENTIRE message — including tool registration. Send tools first,
    // then attempt input_audio_transcription separately so a rejection
    // doesn't break tool registration.
    this.sessionUpdateCount = 0;

    const sessionUpdate = {
      type: 'session.update',
      session: { type: 'realtime', tools }
    };
    this.dc?.send(JSON.stringify(sessionUpdate));
    this.sessionUpdateCount++;
    logService.info('voice', `session.update #${this.sessionUpdateCount} sent: tools (${tools.length} tools)`);

    // Note: input_audio_transcription is not supported by Azure Foundry.
    // User speech transcription is handled via conversation.item.input_audio_transcription
    // events if/when Azure adds support. No second session.update needed.
  }

  private onDataChannelMessage(event: MessageEvent): void {
    let msg: { type: string; [key: string]: unknown };
    try {
      msg = JSON.parse(event.data);
    } catch {
      return;
    }

    switch (msg.type) {
      case 'session.created':
        logService.info('voice', 'Realtime session created');
        break;

      case 'session.updated': {
        // Log what was accepted — check if transcription is enabled
        const session = msg.session as Record<string, unknown> | undefined;
        const hasTranscription = session?.input_audio_transcription !== undefined && session?.input_audio_transcription !== null;
        const toolCount = Array.isArray(session?.tools) ? (session.tools as unknown[]).length : 0;
        logService.info('voice', `Session updated: tools=${toolCount}, transcription=${hasTranscription ? 'enabled' : 'not set'}`);
        break;
      }

      case 'response.created':
        this.responseInFlight = true;
        this.responseStartedAt = Date.now();
        this.queuedResponseCreate = false;
        this.clearQueuedResponseWatchdog();
        this.clearPendingResponseCreatedWatchdog();
        this.resetCurrentResponseState();
        this.setAssistantPlaybackState('buffering');
        logService.info('voice', 'LLM response started');
        this.audioTranscriptBuffer = '';
        break;

      case 'response.done': {
        // Flush any accumulated transcript that wasn't emitted via .done
        if (this.audioTranscriptBuffer) {
          logService.info('voice', `Assistant (flushed): ${this.audioTranscriptBuffer.substring(0, 100)}`);
          this.callbacks?.onTranscript(this.audioTranscriptBuffer, 'assistant');
          this.audioTranscriptBuffer = '';
        }
        const shouldContinueExpressionOnlyResponse = this.shouldContinueExpressionOnlyResponse();
        this.responseInFlight = false;
        this.responseStartedAt = undefined;
        if (!this.outputAudioBufferActive) {
          this.setAssistantPlaybackState('idle');
        }
        logService.info('voice', 'LLM response completed');
        this.resetCurrentResponseState();
        if (shouldContinueExpressionOnlyResponse && !this.queuedResponseCreate) {
          logService.debug('voice', 'Continuing expression-only response after set_expression');
          this.requestResponseCreate('function-output');
          break;
        }
        if (!this.outputAudioBufferActive) {
          this.flushQueuedResponseCreate();
        } else {
          this.scheduleQueuedResponseCreateWatchdog();
        }
        break;
      }

      case 'response.failed':
      case 'response.cancelled':
        this.responseInFlight = false;
        this.responseStartedAt = undefined;
        this.clearPendingResponseCreatedWatchdog();
        this.resetCurrentResponseState();
        this.setAssistantPlaybackState(msg.type === 'response.failed' ? 'error' : 'idle');
        logService.warning('voice', `Response terminated: ${msg.type}`);
        this.flushQueuedResponseCreate();
        break;

      case 'output_audio_buffer.started':
        this.outputAudioBufferActive = true;
        this.outputAudioBufferStartedAt = Date.now();
        this.currentResponseHasUserVisibleOutput = true;
        this.setAssistantPlaybackState('playing');
        logService.debug('voice', 'Event: output_audio_buffer.started');
        break;

      case 'output_audio_buffer.stopped':
        this.outputAudioBufferActive = false;
        this.outputAudioBufferStartedAt = undefined;
        this.setAssistantPlaybackState('idle');
        logService.debug('voice', 'Event: output_audio_buffer.stopped');
        this.flushQueuedResponseCreate();
        break;

      case 'output_audio_buffer.cleared':
        this.outputAudioBufferActive = false;
        this.outputAudioBufferStartedAt = undefined;
        this.setAssistantPlaybackState('idle');
        logService.debug('voice', 'Event: output_audio_buffer.cleared');
        this.flushQueuedResponseCreate();
        break;

      case 'response.audio_transcript.delta':
      case 'response.output_audio_transcript.delta':
      case 'response.output_text.delta':
        // Accumulate streaming transcript chunks (3 possible event naming variants)
        if (typeof msg.delta === 'string') {
          if (msg.delta.length > 0) {
            this.currentResponseHasUserVisibleOutput = true;
          }
          this.audioTranscriptBuffer += msg.delta;
        }
        break;

      case 'response.audio_transcript.done':
      case 'response.output_audio_transcript.done':
      case 'response.output_text.done': {
        // Full transcript — prefer the .done payload, fall back to buffer
        const doneText = (msg.transcript as string) || (msg.text as string) || this.audioTranscriptBuffer;
        if (doneText) {
          this.currentResponseHasUserVisibleOutput = true;
          logService.info('voice', `Assistant: ${doneText.substring(0, 100)}`);
          this.callbacks?.onTranscript(doneText, 'assistant');
        }
        this.audioTranscriptBuffer = '';
        break;
      }

      case 'response.text.done':
        // Legacy text-only response
        if (msg.text) {
          this.currentResponseHasUserVisibleOutput = true;
          this.callbacks?.onTranscript(msg.text as string, 'assistant');
        }
        break;

      case 'conversation.item.input_audio_transcription.completed':
        this.resetVoiceTurnDisplayState();
        logService.info('voice', `User transcript: ${(msg.transcript as string || '').substring(0, 100)}`);
        this.callbacks?.onTranscript(msg.transcript as string, 'user');
        break;

      case 'conversation.item.input_audio_transcription.failed':
        logService.error('voice', `User transcription failed: ${JSON.stringify(msg.error || msg)}`);
        break;

      case 'input_audio_buffer.speech_started':
        this.setState('speaking');
        this.callbacks?.onSpeechActivity?.('started');
        break;

      case 'input_audio_buffer.speech_stopped':
        if (this.state === 'speaking') {
          this.setState('connected');
        }
        this.callbacks?.onSpeechActivity?.('stopped');
        break;

      case 'input_audio_buffer.committed':
        this.resetVoiceTurnDisplayState();
        break;

      case 'response.function_call_arguments.delta':
        this.handleFunctionCallDelta(msg);
        break;

      case 'response.function_call_arguments.done':
        this.deferMessageTask(() => this.handleFunctionCallDone(msg));
        break;

      case 'error': {
        const errorObj = msg.error as { message?: string; type?: string };
        const errorMessage = errorObj?.message || 'Realtime API error';

        // Non-fatal: parameter rejection, missing optional params, or barge-in races.
        const isNonFatal = errorMessage.includes('Unknown parameter')
          || errorMessage.includes('Missing required parameter')
          || /cancellation failed: no active response found/i.test(errorMessage);

        if (isNonFatal) {
          logService.warning('voice', `Non-fatal: ${errorMessage}`);
        } else {
          this.setAssistantPlaybackState('error');
          logService.error('voice', errorMessage);
          this.callbacks?.onError(errorMessage);
        }
        break;
      }

      default:
        logService.debug('voice', `Event: ${msg.type}`);
        break;
    }
  }

  private handleFunctionCallDelta(msg: { call_id?: string; delta?: string; [key: string]: unknown }): void {
    const callId = msg.call_id as string;
    const delta = msg.delta as string;
    if (!callId || !delta) return;

    const existing = this.functionCallBuffers.get(callId) || '';
    this.functionCallBuffers.set(callId, existing + delta);
  }

  private handleFunctionCallDone(msg: { call_id?: string; name?: string; arguments?: string; [key: string]: unknown }): void {
    const callId = msg.call_id as string;
    const funcName = msg.name as string;
    const argsStr = (msg.arguments as string) || this.functionCallBuffers.get(callId || '') || '';
    this.functionCallBuffers.delete(callId || '');

    logService.info('voice', `Function call: ${funcName}`, argsStr.substring(0, 200));
    if (funcName) {
      this.currentResponseToolCalls.push(funcName);
    }

    let args: Record<string, unknown>;
    try {
      args = JSON.parse(argsStr);
    } catch {
      this.sendFunctionOutput(callId, JSON.stringify({
        success: false,
        error: `Invalid JSON arguments for ${funcName}`
      }), funcName);
      return;
    }

    if (shouldSuppressRealtimeDisplayToolCall(funcName, this.displayToolsShownInTurn)) {
      logService.info('voice', `Suppressed redundant display tool: ${funcName}`);
      this.sendFunctionOutput(callId, JSON.stringify({
        success: true,
        alreadyDisplayed: true,
        message: 'Data is already displayed to the user in the action panel.',
        advice: 'Do not create a duplicate display block. Comment on or refine the existing results instead.'
      }), funcName);
      return;
    }

    // Delegate all function calls to the handler
    if (this.callbacks?.onFunctionCall) {
      Promise.resolve()
        .then(() => this.callbacks?.onFunctionCall(callId, funcName, args))
        .then((outputStr) => {
          if (typeof outputStr !== 'string') {
            throw new Error(`Tool ${funcName} did not return a string output`);
          }
          if (isSuccessfulFunctionOutput(outputStr) && VOICE_DUPLICATE_DISPLAY_TOOLS.has(funcName)) {
            this.displayToolsShownInTurn.add(funcName);
          }
          this.sendFunctionOutput(callId, outputStr, funcName);
        })
        .catch((error: unknown) => {
          const normalizedError = normalizeError(error, `Tool ${funcName} failed`);
          logService.error('voice', normalizedError.message);
          this.sendFunctionOutput(callId, JSON.stringify({
            success: false,
            error: normalizedError.message
          }), funcName);
        });
    }
  }

  /**
   * Send function call output back to the model and trigger next response.
   */
  private sendFunctionOutput(callId: string, output: string, funcName?: string): void {
    if (this.dc && callId) {
      const msg = {
        type: 'conversation.item.create',
        item: {
          type: 'function_call_output',
          call_id: callId,
          output
        }
      };
      this.dc.send(JSON.stringify(msg));
      if (funcName === 'set_expression') {
        this.pendingExpressionOnlyFollowUp = true;
      }
      if (shouldDeferImmediateResponse(funcName, output)) {
        logService.debug('voice', `Deferred immediate response.create for ${funcName} (awaiting final tool context)`);
      } else {
        this.requestResponseCreate('function-output');
      }
    }
  }

  private requestResponseCreate(reason: ResponseCreateReason, responsePayload?: Record<string, unknown>): void {
    if (!this.dc || this.dc.readyState !== 'open') return;

    const now = Date.now();

    if (this.responseInFlight || this.outputAudioBufferActive || (now - this.lastResponseCreateAt) < RESPONSE_CREATE_RECENT_WINDOW_MS) {
      if (responsePayload) {
        this.queuedResponseCreatePayload = responsePayload;
      }
      if (!this.queuedResponseCreate) {
        logService.debug('voice', `Queued response.create (${reason})`);
      }
      this.queuedResponseCreate = true;
      this.scheduleQueuedResponseCreateWatchdog();
      return;
    }

    this.lastResponseCreateAt = now;
    const message = responsePayload
      ? { type: 'response.create', response: responsePayload }
      : { type: 'response.create' };
    if (responsePayload) {
      this.queuedResponseCreatePayload = undefined;
    }
    this.dc.send(JSON.stringify(message));
    this.armPendingResponseCreatedWatchdog();
  }

  private flushQueuedResponseCreate(): void {
    if (!this.queuedResponseCreate) return;
    if (this.outputAudioBufferActive) return;
    this.queuedResponseCreate = false;
    this.clearQueuedResponseWatchdog();
    const responsePayload = this.queuedResponseCreatePayload;
    this.requestResponseCreate('queued', responsePayload);
  }

  private scheduleQueuedResponseCreateWatchdog(): void {
    this.clearQueuedResponseWatchdog();
    this.queuedResponseWatchdog = setTimeout(() => {
      this.queuedResponseWatchdog = undefined;

      if (!this.queuedResponseCreate || !this.dc || this.dc.readyState !== 'open') {
        return;
      }

      if (this.responseInFlight) {
        this.scheduleQueuedResponseCreateWatchdog();
        return;
      }

      if (this.outputAudioBufferActive) {
        const outputAudioAgeMs = this.outputAudioBufferStartedAt
          ? (Date.now() - this.outputAudioBufferStartedAt)
          : 0;
        if (outputAudioAgeMs < OUTPUT_AUDIO_BUFFER_STALL_MS) {
          this.scheduleQueuedResponseCreateWatchdog();
          return;
        }

        logService.warning('voice', 'Queued response.create watchdog cleared stale output_audio_buffer state');
        this.outputAudioBufferActive = false;
        this.outputAudioBufferStartedAt = undefined;
      }

      logService.warning('voice', 'Queued response.create watchdog flushed a stalled response queue');
      this.flushQueuedResponseCreate();
    }, QUEUED_RESPONSE_WATCHDOG_MS);
  }

  private clearQueuedResponseWatchdog(): void {
    if (!this.queuedResponseWatchdog) return;
    clearTimeout(this.queuedResponseWatchdog);
    this.queuedResponseWatchdog = undefined;
  }

  private armPendingResponseCreatedWatchdog(): void {
    this.clearPendingResponseCreatedWatchdog();
    this.pendingResponseCreatedWatchdog = setTimeout(() => {
      this.pendingResponseCreatedWatchdog = undefined;

      if (this.responseInFlight || !this.dc || this.dc.readyState !== 'open') {
        return;
      }

      logService.warning('voice', 'response.create watchdog fired before response.created; retrying once');
      this.lastResponseCreateAt = 0;
      this.requestResponseCreate('queued');
    }, RESPONSE_CREATED_WATCHDOG_MS);
  }

  private clearPendingResponseCreatedWatchdog(): void {
    if (!this.pendingResponseCreatedWatchdog) return;
    clearTimeout(this.pendingResponseCreatedWatchdog);
    this.pendingResponseCreatedWatchdog = undefined;
  }

  private deferMessageTask(task: () => void): void {
    const taskId = setTimeout(() => {
      this.deferredMessageTaskIds = this.deferredMessageTaskIds.filter((id) => id !== taskId);
      task();
    }, 0);
    this.deferredMessageTaskIds.push(taskId);
  }

  private clearDeferredMessageTasks(): void {
    for (const taskId of this.deferredMessageTaskIds) {
      clearTimeout(taskId);
    }
    this.deferredMessageTaskIds = [];
  }

  private handleChannelFailure(message: string): void {
    if (this.state === 'idle' || this.state === 'error') return;
    logService.warning('voice', message);
    this.callbacks?.onError(message);
    this.cleanup();
    this.setAssistantPlaybackState('error');
    this.setState('error');
  }

  private handleChannelError(): void {
    const peerState = this.pc?.connectionState;
    const channelState = this.dc?.readyState;
    logService.warning('voice', 'Realtime data channel error');

    if (peerState === 'failed' || peerState === 'closed' || channelState === 'closed') {
      this.handleChannelFailure('Realtime data channel error');
    }
  }

  private setState(state: RealtimeState): void {
    this.state = state;
    this.callbacks?.onStateChange(state);
  }

  private setAssistantPlaybackState(state: AssistantPlaybackState): void {
    if (this.assistantPlaybackState === state) return;
    this.assistantPlaybackState = state;
    this.callbacks?.onAssistantPlaybackStateChange?.(state);
  }

  private async fetchWithTimeout(
    url: string,
    init: RequestInit,
    timeoutMs: number,
    operation: string
  ): Promise<Response> {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);

    try {
      return await fetch(url, { ...init, signal: controller.signal });
    } catch (error) {
      const normalizedError = normalizeError(error, `${operation} failed`);
      if (normalizedError.name === 'AbortError') {
        throw new Error(`${operation} timed out after ${Math.round(timeoutMs / 1000)}s`);
      }
      throw error;
    } finally {
      clearTimeout(timeout);
    }
  }

  private cleanup(): void {
    if (this.mediaStream) {
      this.mediaStream.getTracks().forEach((track) => track.stop());
      this.mediaStream = undefined;
    }

    if (this.dc) {
      this.dc.onopen = null;
      this.dc.onmessage = null;
      this.dc.onerror = null;
      this.dc.onclose = null;
      this.dc.close();
      this.dc = undefined;
    }

    if (this.pc) {
      this.pc.ontrack = null;
      this.pc.onicecandidate = null;
      this.pc.onconnectionstatechange = null;
      this.pc.oniceconnectionstatechange = null;
      this.pc.ondatachannel = null;
      this.pc.close();
      this.pc = undefined;
    }

    if (this.audioElement) {
      this.audioElement.onended = null;
      this.audioElement.onerror = null;
      this.audioElement.onpause = null;
      this.audioElement.onplay = null;
      this.audioElement.srcObject = null;
      this.audioElement = undefined;
    }

    this.remoteStream = undefined;
    this.clearQueuedResponseWatchdog();
    this.clearPendingResponseCreatedWatchdog();
    this.clearDeferredMessageTasks();
    this.functionCallBuffers.clear();
    this.responseInFlight = false;
    this.queuedResponseCreate = false;
    this.lastResponseCreateAt = 0;
    this.outputAudioBufferActive = false;
    this.outputAudioBufferStartedAt = undefined;
    this.responseStartedAt = undefined;
    this.displayToolsShownInTurn.clear();
    this.resetCurrentResponseState();
    this.setAssistantPlaybackState('idle');
    this.callbacks?.onRemoteStream?.(undefined);
  }

  private resetVoiceTurnDisplayState(): void {
    this.displayToolsShownInTurn.clear();
  }

  private resetCurrentResponseState(): void {
    this.currentResponseHasUserVisibleOutput = false;
    this.currentResponseToolCalls = [];
    this.pendingExpressionOnlyFollowUp = false;
  }

  private shouldContinueExpressionOnlyResponse(): boolean {
    return this.pendingExpressionOnlyFollowUp
      && !this.currentResponseHasUserVisibleOutput
      && this.currentResponseToolCalls.length === 1
      && this.currentResponseToolCalls[0] === 'set_expression';
  }
}

export function shouldDeferImmediateResponse(funcName: string | undefined, output: string): boolean {
  if (!funcName || !DEFERRED_RESPONSE_TOOLS.has(funcName)) return false;

  if (funcName === 'show_compose_form') {
    return true;
  }

  if (funcName === 'set_expression') {
    return true;
  }

  try {
    const parsed = JSON.parse(output) as Record<string, unknown>;
    if (parsed.success !== true) return false;
    if (parsed.workflowHandled === true) return true;

    const message = typeof parsed.message === 'string' ? parsed.message.trim().toLowerCase() : '';
    if (!message) return true;

    return (
      message.includes('results will appear')
      || message.includes('will appear in the panel')
      || message.startsWith('searching ')
      || message.startsWith('executing ')
      || message.includes('hold on')
      || message.includes('shortly')
    );
  } catch {
    return false;
  }
}

function isSuccessfulFunctionOutput(output: string): boolean {
  try {
    const parsed = JSON.parse(output) as Record<string, unknown>;
    return parsed.success === true;
  } catch {
    return false;
  }
}

export function shouldSuppressRealtimeDisplayToolCall(
  toolName: string | undefined,
  displayToolsShownInTurn: ReadonlySet<string>
): boolean {
  return !!toolName
    && VOICE_DUPLICATE_DISPLAY_TOOLS.has(toolName)
    && displayToolsShownInTurn.has(toolName);
}

// Export singleton
export const realtimeAudioService = new RealtimeAudioService();
