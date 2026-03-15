/**
 * useVoiceSession
 * Hook that encapsulates the voice session lifecycle:
 * connect, disconnect, mute, health check, function call dispatch,
 * and expression auto-triggers from sentiment analysis.
 *
 * Dual-path text: when voice is connected, text goes through WebRTC.
 * When voice is NOT connected, text goes through HTTP chat completions
 * via TextChatService — same tool dispatch, same UI blocks.
 */

import * as React from 'react';
import { realtimeAudioService } from '../services/realtime/RealtimeAudioService';
import type { RealtimeState } from '../services/realtime/RealtimeAudioService';
import { TextChatService } from '../services/realtime/TextChatService';
import type { ITextChatCallbacks } from '../services/realtime/TextChatService';
import { getSystemPrompt } from '../services/realtime/SystemPrompt';
import type { IPromptConfig } from '../services/realtime/SystemPrompt';
import { getTools } from '../services/realtime/ToolRegistry';
import { useGrimoireStore } from '../store/useGrimoireStore';
import type { ConnectionState, ITranscriptEntry } from '../store/useGrimoireStore';
import type { Expression } from '../services/avatar/ExpressionEngine';
import { analyzeSentiment, analyzeSentimentAsync, isQuestion } from '../services/avatar/SentimentAnalyzer';
import { getNanoService } from '../services/nano/NanoService';
import type { NanoService } from '../services/nano/NanoService';
import { logService } from '../services/logging/LogService';
import { hybridInteractionEngine } from '../services/hie/HybridInteractionEngine';
import { createCorrelationId } from '../services/hie/HAEContracts';
import { isHieToolErrorPromptMessage } from '../services/hie/HiePromptProtocol';
import { resolveIngressTurnStartPolicy } from '../services/hie/HieTurnStartPolicy';
import { handleFunctionCall } from '../services/tools/handleFunctionCall';
import { normalizeRealtimeVoiceId } from '../services/realtime/RealtimeVoiceCatalog';
import { getToolAckText } from '../services/tools/ToolAcknowledgment';
import { executeCompoundWorkflowPlan, type ICompoundWorkflowPlan } from '../services/realtime/CompoundWorkflowExecutor';
import {
  getToolCompletionAckText,
  getVoiceToolCompletionAckText,
  shouldUseImmediateCompletionAck
} from '../services/tools/ToolCompletionAcknowledgment';
import { getUserVisibleTextChatError } from '../services/realtime/TextChatFeedback';
import { createRateLimitLifecycleCallbacks } from '../services/realtime/TextChatRateLimitUi';
import {
  buildConversationLanguageContextMessage,
  resolveConversationLanguage
} from '../services/context/ConversationLanguage';
import {
  buildFirstTurnRoutingLogDetail,
  classifyAssistantFirstTurnOutcome,
  getFirstTurnRoutingObservation,
  getObservedFirstToolName,
  type IFirstTurnRoutingObservation
} from '../services/realtime/IntentRoutingPolicy';
import {
  resolveMailDiscussionReplyAllPlan,
  resolveMailDiscussionReplyAllPlanFromToolCall
} from '../services/realtime/MailDiscussionReplyAllWorkflow';
import { fetchBackendHealth } from '../services/startup/BackendHealthService';

// ─── Expression with Auto-Revert ─────────────────────────────

function useExpressionControl(): {
  setExpressionWithRevert: (expr: Expression, revertMs?: number) => void;
} {
  const setExpression = useGrimoireStore((s) => s.setExpression);
  const revertTimerRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  const setExpressionWithRevert = React.useCallback((expr: Expression, revertMs?: number) => {
    if (revertTimerRef.current) {
      clearTimeout(revertTimerRef.current);
      revertTimerRef.current = undefined;
    }
    setExpression(expr);
    if (revertMs && revertMs > 0) {
      revertTimerRef.current = setTimeout(() => {
        setExpression('idle');
        revertTimerRef.current = undefined;
      }, revertMs);
    }
  }, [setExpression]);

  React.useEffect(() => {
    return () => {
      if (revertTimerRef.current) {
        clearTimeout(revertTimerRef.current);
      }
    };
  }, []);

  return { setExpressionWithRevert };
}

function createTranscriptEntry(
  text: string,
  role: ITranscriptEntry['role'],
  turnId?: string
): ITranscriptEntry {
  const lineage = hybridInteractionEngine.getTurnLineage(turnId);
  return {
    text,
    role,
    timestamp: new Date(),
    turnId,
    rootTurnId: lineage?.rootTurnId || turnId,
    parentTurnId: lineage?.parentTurnId
  };
}

function getLatestUserTranscriptText(): string {
  const transcript = useGrimoireStore.getState().transcript;
  for (let i = transcript.length - 1; i >= 0; i--) {
    if (transcript[i].role === 'user') return transcript[i].text;
  }
  return '';
}

function getLatestUserTranscriptTextForTurn(turnId: string | undefined): string {
  if (!turnId) {
    return '';
  }

  const transcript = useGrimoireStore.getState().transcript;
  for (let i = transcript.length - 1; i >= 0; i--) {
    if (transcript[i].role !== 'user') {
      continue;
    }
    if (transcript[i].turnId === turnId && transcript[i].text.trim()) {
      return transcript[i].text;
    }
  }

  return '';
}

function beginUserTurnWithIngressPolicy(text: string): string {
  const snapshot = hybridInteractionEngine.getVisualStateSnapshot();
  const turnPolicy = resolveIngressTurnStartPolicy(text, {
    hasTaskContext: !!hybridInteractionEngine.getCurrentTaskContext(),
    hasVisibleBlocks: snapshot.blocks.length > 0,
    visibleBlockTitles: snapshot.blocks.map((block) => block.title),
    visibleReferenceTitles: snapshot.referencesNewestFirst.flatMap((entry) => (
      entry.references.map((reference) => reference.title).filter((title) => title.trim().length > 0)
    ))
  });

  return hybridInteractionEngine.beginUserTurn({
    turnId: createCorrelationId('turn'),
    mode: turnPolicy.mode,
    reason: turnPolicy.reason,
    text
  }).turnId;
}

// ─── Streaming Callbacks Factory ─────────────────────────────────

function createStreamingCallbacks(
  setExpressionWithRevert: (expr: Expression, revertMs?: number) => void,
  label: string,
  onAssistantComplete?: (text: string) => void,
  options?: {
    surfaceErrors?: boolean;
    resolveTurnId?: () => string | undefined;
  }
): ITextChatCallbacks {
  let streamingStarted = false;
  let toolLoopActive = false;
  let pendingTokenChunk = '';
  let flushTimer: ReturnType<typeof setTimeout> | undefined;
  const TOKEN_FLUSH_MS = 50;
  const TOOL_ACK_DEDUPE_MS = 1500;
  let localAckSent = false;
  let lastLocalAckAt = 0;
  const rateLimitLifecycle = createRateLimitLifecycleCallbacks({
    resolveTurnId: options?.resolveTurnId
  });

  const emitLocalToolAck = (funcName: string, args: Record<string, unknown>): void => {
    const ack = getToolAckText(funcName, args, useGrimoireStore.getState().conversationLanguage);
    if (!ack) return;
    const now = Date.now();
    if (localAckSent || (now - lastLocalAckAt) < TOOL_ACK_DEDUPE_MS) return;
    localAckSent = true;
    lastLocalAckAt = now;
    useGrimoireStore.getState().addTranscript(createTranscriptEntry(ack, 'assistant', options?.resolveTurnId?.()));
  };

  const flushPendingTokenChunk = (): void => {
    if (!pendingTokenChunk) return;
    const chunk = pendingTokenChunk;
    pendingTokenChunk = '';

    const s = useGrimoireStore.getState();
    if (!streamingStarted) {
      s.setActivityStatus('');
      s.addTranscript(createTranscriptEntry(chunk, 'assistant', options?.resolveTurnId?.()));
      streamingStarted = true;
      return;
    }

    const last = s.transcript[s.transcript.length - 1];
    if (last && last.role === 'assistant') {
      s.updateLastTranscript(last.text + chunk);
    } else {
      s.addTranscript(createTranscriptEntry(chunk, 'assistant', options?.resolveTurnId?.()));
    }
  };

  const scheduleFlush = (): void => {
    if (flushTimer) return;
    flushTimer = setTimeout(() => {
      flushTimer = undefined;
      flushPendingTokenChunk();
    }, TOKEN_FLUSH_MS);
  };

  return {
    onToken: (chunk: string) => {
      pendingTokenChunk += chunk;
      scheduleFlush();
    },
    onFunctionCall: (_callId: string, funcName: string, args: Record<string, unknown>): string | Promise<string> => {
      if (flushTimer) {
        clearTimeout(flushTimer);
        flushTimer = undefined;
      }
      flushPendingTokenChunk();
      emitLocalToolAck(funcName, args);
      // Suppress expression reverts during multi-tool loops to prevent flashing
      if (!toolLoopActive) {
        toolLoopActive = true;
        hybridInteractionEngine.suppressExpressionReverts();
      }
      return handleFunctionCall(funcName, args, useGrimoireStore.getState(), true);
    },
    onComplete: (fullText: string) => {
      if (flushTimer) {
        clearTimeout(flushTimer);
        flushTimer = undefined;
      }
      flushPendingTokenChunk();
      if (toolLoopActive) {
        hybridInteractionEngine.releaseExpressionReverts();
        toolLoopActive = false;
      }
      rateLimitLifecycle.clear();
      const s = useGrimoireStore.getState();
      if (fullText) {
        const last = s.transcript[s.transcript.length - 1];
        if (last && last.role === 'assistant') {
          s.updateLastTranscript(fullText);
        } else {
          s.addTranscript(createTranscriptEntry(fullText, 'assistant', options?.resolveTurnId?.()));
        }
      }
      s.setActivityStatus('');
      s.setTextChatActive(false);
      hybridInteractionEngine.onLlmResponse();
      if (fullText && onAssistantComplete) {
        onAssistantComplete(fullText);
      }
      setExpressionWithRevert('happy', 2000);
      logService.info('llm', `${label}: "${fullText.substring(0, 80)}..."`);
    },
    onError: (error: string) => {
      if (flushTimer) {
        clearTimeout(flushTimer);
        flushTimer = undefined;
      }
      pendingTokenChunk = '';
      if (toolLoopActive) {
        hybridInteractionEngine.releaseExpressionReverts();
        toolLoopActive = false;
      }
      const s = useGrimoireStore.getState();
      if (!rateLimitLifecycle.isExhausted()) {
        rateLimitLifecycle.clear();
      }
      s.setTextChatActive(false);
      if (options?.surfaceErrors) {
        const visibleError = getUserVisibleTextChatError(error);
        const last = s.transcript[s.transcript.length - 1];
        if ((streamingStarted || localAckSent) && last && last.role === 'assistant') {
          s.updateLastTranscript(visibleError);
        } else {
          s.addTranscript(createTranscriptEntry(visibleError, 'assistant', options?.resolveTurnId?.()));
        }
      }
      setExpressionWithRevert('confused', 3000);
      logService.error('llm', `${label} error: ${error}`);
    },
    onRateLimitRetry: rateLimitLifecycle.onRateLimitRetry,
    onRateLimitResolved: rateLimitLifecycle.onRateLimitResolved,
    onRateLimitExhausted: rateLimitLifecycle.onRateLimitExhausted
  };
}

// ─── Hook ───────────────────────────────────────────────────────

export interface IUseVoiceSession {
  connect: () => Promise<void>;
  reconnect: () => Promise<void>;
  disconnect: () => void;
  toggleMute: () => void;
  sendText: (text: string) => void;
  checkHealth: () => Promise<void>;
}

export function useVoiceSession(): IUseVoiceSession {
  const proxyConfig = useGrimoireStore((s) => s.proxyConfig);
  const avatarEnabled = useGrimoireStore((s) => s.avatarEnabled);
  const personality = useGrimoireStore((s) => s.personality);
  const visage = useGrimoireStore((s) => s.visage);
  const userContext = useGrimoireStore((s) => s.userContext);
  const isMuted = useGrimoireStore((s) => s.isMuted);
  const setConnectionState = useGrimoireStore((s) => s.setConnectionState);
  const setRemoteStream = useGrimoireStore((s) => s.setRemoteStream);
  const setMuted = useGrimoireStore((s) => s.setMuted);
  const setMicGranted = useGrimoireStore((s) => s.setMicGranted);
  const setHealthCheckState = useGrimoireStore((s) => s.setHealthCheckState);
  const setAssistantPlaybackState = useGrimoireStore((s) => s.setAssistantPlaybackState);
  const addTranscript = useGrimoireStore((s) => s.addTranscript);
  const setExpression = useGrimoireStore((s) => s.setExpression);
  const conversationLanguage = useGrimoireStore((s) => s.conversationLanguage);
  const setConversationLanguage = useGrimoireStore((s) => s.setConversationLanguage);

  const { setExpressionWithRevert } = useExpressionControl();
  const nanoServiceRef = React.useRef<NanoService | undefined>(undefined);
  const textChatServiceRef = React.useRef<TextChatService | undefined>(undefined);
  const voiceAckStateRef = React.useRef<{ sentInCycle: boolean; lastAckAt: number }>({
    sentInCycle: false,
    lastAckAt: 0
  });
  const voiceCompletionAckStateRef = React.useRef<{ lastAckAt: number; lastSignature: string }>({
    lastAckAt: 0,
    lastSignature: ''
  });
  const voiceTurnRoutingRef = React.useRef<IFirstTurnRoutingObservation | undefined>(undefined);
  const activeTurnIdRef = React.useRef<string | undefined>(undefined);
  const mailDiscussionWorkflowHandledTurnRef = React.useRef<string | undefined>(undefined);

  const updateConversationLanguageForTurn = React.useCallback((text: string): string => {
    const currentStore = useGrimoireStore.getState();
    const nextLanguage = resolveConversationLanguage(
      text,
      currentStore.conversationLanguage,
      currentStore.userContext?.resolvedLanguage
    );
    if (nextLanguage !== currentStore.conversationLanguage) {
      currentStore.setConversationLanguage(nextLanguage);
      if (realtimeAudioService.isConnected() && (currentStore.conversationLanguage || nextLanguage !== 'en')) {
        realtimeAudioService.sendContextMessage(buildConversationLanguageContextMessage(nextLanguage), false);
      }
    }
    return nextLanguage;
  }, []);

  const resetVoiceAckCycle = React.useCallback((): void => {
    voiceAckStateRef.current.sentInCycle = false;
    voiceCompletionAckStateRef.current.lastSignature = '';
  }, []);

  const maybeEmitVoiceToolAck = React.useCallback((funcName: string, args: Record<string, unknown>): void => {
    const ack = getToolAckText(funcName, args, useGrimoireStore.getState().conversationLanguage);
    if (!ack) return;
    const now = Date.now();
    const state = voiceAckStateRef.current;
    if (state.sentInCycle || (now - state.lastAckAt) < 1500) return;
    state.sentInCycle = true;
    state.lastAckAt = now;
    const spokeViaRealtime = realtimeAudioService.speakDeterministicText(ack);
    if (!spokeViaRealtime) {
      useGrimoireStore.getState().addTranscript(
        createTranscriptEntry(ack, 'assistant', activeTurnIdRef.current || hybridInteractionEngine.getCurrentTurnId())
      );
    }
  }, []);

  const maybeEmitVoiceToolCompletionAck = React.useCallback((toolName: string, itemCount: number): void => {
    const ack = getToolCompletionAckText(toolName, itemCount, useGrimoireStore.getState().conversationLanguage);
    if (!ack) return;

    const state = voiceCompletionAckStateRef.current;
    const now = Date.now();
    const signature = `${toolName}:${itemCount}:${ack}`;
    if (state.lastSignature === signature && (now - state.lastAckAt) < 2000) return;

    state.lastAckAt = now;
    state.lastSignature = signature;
    const spokeViaRealtime = shouldUseImmediateCompletionAck(toolName, getLatestUserTranscriptText())
      && realtimeAudioService.speakDeterministicText(ack);
    if (!spokeViaRealtime) {
      useGrimoireStore.getState().addTranscript(
        createTranscriptEntry(ack, 'assistant', activeTurnIdRef.current || hybridInteractionEngine.getCurrentTurnId())
      );
    }
  }, []);

  const maybeEmitVoiceImmediateToolResultAck = React.useCallback((
    toolName: string,
    args: Record<string, unknown>,
    output: string
  ): void => {
    const ack = getVoiceToolCompletionAckText(toolName, args, output, useGrimoireStore.getState().conversationLanguage);
    if (!ack) return;

    const state = voiceCompletionAckStateRef.current;
    const now = Date.now();
    const signature = `${toolName}:${ack}`;
    if (state.lastSignature === signature && (now - state.lastAckAt) < 2000) return;

    state.lastAckAt = now;
    state.lastSignature = signature;
    const spokeViaRealtime = toolName === 'show_compose_form'
      && realtimeAudioService.speakDeterministicText(ack);
    if (!spokeViaRealtime) {
      useGrimoireStore.getState().addTranscript(
        createTranscriptEntry(ack, 'assistant', activeTurnIdRef.current || hybridInteractionEngine.getCurrentTurnId())
      );
    }
  }, []);

  const beginVoiceTurnRoutingObservation = React.useCallback((text: string): void => {
    voiceTurnRoutingRef.current = getFirstTurnRoutingObservation(text);
  }, []);

  const logVoiceTurnRoutingOutcome = React.useCallback((
    outcome: 'tool_call' | 'clarification' | 'answer_only',
    actualToolName?: string
  ): void => {
    const observation = voiceTurnRoutingRef.current;
    if (!observation) return;

    const detail = buildFirstTurnRoutingLogDetail('voice', outcome, observation, actualToolName);
    logService.info('llm', `First-turn routing: ${outcome}`, detail);
    if (outcome === 'clarification' && observation.isGenericEnterpriseSearch) {
      logService.warning('llm', 'Generic enterprise search fell back to clarification', detail);
    }

    voiceTurnRoutingRef.current = undefined;
  }, []);

  const emitDeterministicAssistantText = React.useCallback((text: string): void => {
    const trimmed = text.trim();
    if (!trimmed) return;

    const spokeViaRealtime = realtimeAudioService.speakDeterministicText(trimmed);
    if (!spokeViaRealtime) {
      useGrimoireStore.getState().addTranscript(
        createTranscriptEntry(trimmed, 'assistant', activeTurnIdRef.current || hybridInteractionEngine.getCurrentTurnId())
      );
    }
  }, []);

  const runVoiceCompoundWorkflow = React.useCallback(async (
    plan: ICompoundWorkflowPlan,
    options?: { speakAssistantText?: boolean }
  ): Promise<string> => {
    const assistantText = await executeCompoundWorkflowPlan(plan, {
      onFunctionCall: (_callId: string, funcName: string, args: Record<string, unknown>) => (
        handleFunctionCall(funcName, args, useGrimoireStore.getState(), true)
      )
    });
    const activeTurnId = activeTurnIdRef.current || hybridInteractionEngine.getCurrentTurnId();
    if (activeTurnId) {
      mailDiscussionWorkflowHandledTurnRef.current = activeTurnId;
    }
    if (options?.speakAssistantText) {
      emitDeterministicAssistantText(assistantText);
    }
    return assistantText;
  }, [emitDeterministicAssistantText]);

  // Initialize TextChatService when prompt inputs become available.
  // Re-creates the service when user context or visage changes so the system prompt stays in sync.
  React.useEffect(() => {
    if (proxyConfig) {
      const state = useGrimoireStore.getState();
      const promptConfig: IPromptConfig = {
        mcpEnvironmentId: state.mcpEnvironmentId,
        hasGraphAccess: !!state.aadHttpClient,
        avatarEnabled,
        conversationLanguage: state.conversationLanguage
      };
      textChatServiceRef.current = new TextChatService(
        proxyConfig,
        personality,
        userContext,
        promptConfig,
        visage
      );
      logService.debug('system', 'TextChatService initialized');
    } else {
      textChatServiceRef.current = undefined;
    }
  }, [proxyConfig, personality, userContext, avatarEnabled, visage, conversationLanguage]);

  const connect = React.useCallback(async () => {
    if (!proxyConfig) {
      logService.error('system', 'No proxy config — cannot connect');
      return;
    }

    const currentState = useGrimoireStore.getState();
    const voice = normalizeRealtimeVoiceId(currentState.voiceId);
    if (voice !== currentState.voiceId) {
      currentState.setVoiceId(voice);
    }
    const instructions = getSystemPrompt(personality, currentState.visage, currentState.userContext, {
      mcpEnvironmentId: currentState.mcpEnvironmentId,
      hasGraphAccess: !!currentState.aadHttpClient,
      avatarEnabled: currentState.avatarEnabled,
      conversationLanguage: currentState.conversationLanguage
    });
    const tools = getTools({ avatarEnabled: currentState.avatarEnabled });

    logService.info('voice', `Connecting with voice="${voice}", personality="${personality}", visage="${visage}"`);
    setExpression('thinking');

    await realtimeAudioService.connect(
      proxyConfig,
      {
        onStateChange: (state: RealtimeState) => {
          // Map RealtimeState to ConnectionState
          const mapped: ConnectionState = state as ConnectionState;
          setConnectionState(mapped);

          if (state === 'connected') {
            hybridInteractionEngine.setVoicePathActive(true);
            hybridInteractionEngine.setAsyncToolCompletionHandler((feedback) => {
              maybeEmitVoiceToolCompletionAck(feedback.toolName, feedback.itemCount);
            });
            setMicGranted(true);
            setAssistantPlaybackState('idle');

            // Always capture streams (may change on reconnect)
            const stream = realtimeAudioService.getRemoteStream();
            if (stream) {
              setRemoteStream(stream);
            }
            const micStream = realtimeAudioService.getMicStream();
            if (micStream) {
              useGrimoireStore.getState().setMicStream(micStream);
            }

            // Only initialize HIE once per session
            if (!hybridInteractionEngine.isInitialized()) {
              setExpression('happy');
              logService.info('voice', 'Voice session connected — initializing HIE');

              // Initialize NanoService if fast backend is available
              nanoServiceRef.current = getNanoService(proxyConfig);

              // Initialize HIE engine (with Nano for smart compression)
              const setGazeTarget = useGrimoireStore.getState().setGazeTarget;
              hybridInteractionEngine.initialize(setExpression, {
                setGazeFn: setGazeTarget,
                nanoService: nanoServiceRef.current,
                sendContextMessage: (text, trigger, turnId) => {
                  if (turnId) {
                    activeTurnIdRef.current = turnId;
                    hybridInteractionEngine.setCurrentTurnId(turnId);
                  }
                  if (trigger) {
                    const pendingWorkflowPlan = resolveMailDiscussionReplyAllPlan(undefined, {
                      allowPending: true
                    });
                    if (pendingWorkflowPlan) {
                      runVoiceCompoundWorkflow(pendingWorkflowPlan, { speakAssistantText: true }).catch((error: Error) => {
                        logService.error('voice', `Pending mail discussion workflow failed: ${error.message}`);
                        realtimeAudioService.sendContextMessage(text, trigger);
                      });
                      return;
                    }
                  }
                  realtimeAudioService.sendContextMessage(text, trigger);
                },
                onAsyncToolCompletion: (feedback) => {
                  maybeEmitVoiceToolCompletionAck(feedback.toolName, feedback.itemCount);
                }
              });

              // Revert to idle after 1.5s
              setTimeout(() => setExpression('idle'), 1500);
            }
          }
        },
        onAssistantPlaybackStateChange: (state) => {
          setAssistantPlaybackState(state);
        },
        onRemoteStream: (stream) => {
          setRemoteStream(stream);
        },
        onTranscript: (text: string, role: 'user' | 'assistant') => {
          const turnId = role === 'user'
            ? beginUserTurnWithIngressPolicy(text)
            : (activeTurnIdRef.current || hybridInteractionEngine.getCurrentTurnId());
          if (turnId) {
            activeTurnIdRef.current = turnId;
            hybridInteractionEngine.setCurrentTurnId(turnId);
          }
          const entry: ITranscriptEntry = createTranscriptEntry(text, role, turnId);
          addTranscript(entry);
          if (role === 'user') {
            mailDiscussionWorkflowHandledTurnRef.current = undefined;
            updateConversationLanguageForTurn(text);
            resetVoiceAckCycle();
            beginVoiceTurnRoutingObservation(text);
          }

          // Expression auto-triggers from sentiment (sync regex first)
          const sentiment = analyzeSentiment(text, role);
          if (sentiment.expression) {
            setExpressionWithRevert(sentiment.expression, sentiment.revertMs);
          } else if (role === 'user' && isQuestion(text)) {
            setExpressionWithRevert('listening', 2000);
          }

          // Async Nano enhancement (fire-and-forget if regex found nothing)
          if (!sentiment.expression && nanoServiceRef.current) {
            analyzeSentimentAsync(text, role, nanoServiceRef.current).then((nanoResult) => {
              if (nanoResult.expression) {
                setExpressionWithRevert(nanoResult.expression, nanoResult.revertMs);
              }
            }).catch(() => { /* Nano sentiment is best-effort */ });
          }

          // Dynamic visual state grounding on LLM responses
          if (role === 'assistant') {
            logVoiceTurnRoutingOutcome(classifyAssistantFirstTurnOutcome(text));
            hybridInteractionEngine.onLlmResponse();
            useGrimoireStore.getState().setActivityStatus('');
          }

          logService.info('voice', `Transcript [${role}]: ${text.substring(0, 80)}`);
        },
        onSpeechActivity: (activity: 'started' | 'stopped') => {
          const currentStore = useGrimoireStore.getState();
          if (activity === 'started') {
            realtimeAudioService.interruptAssistantPlayback();
            currentStore.setActivityStatus('');
            resetVoiceAckCycle();
            logService.debug('voice', 'User speech started');
          } else {
            logService.debug('voice', 'User speech stopped — awaiting response');
          }
        },
        onError: (error: string) => {
          setAssistantPlaybackState('error');
          logService.error('voice', error);
          setExpressionWithRevert('confused', 3000);
        },
        onFunctionCall: async (_callId: string, funcName: string, args: Record<string, unknown>): Promise<string> => {
          // Get fresh store state for each function call
          const currentStore = useGrimoireStore.getState();
          // Guard: ensure a turn exists before tool execution.
          // In voice mode, GPT Realtime API may fire onFunctionCall before onTranscript.
          if (!activeTurnIdRef.current) {
            const fallbackLineage = hybridInteractionEngine.beginUserTurn({
              turnId: createCorrelationId('turn'),
              mode: 'inherit',
              reason: 'voice-tool-before-transcript'
            });
            activeTurnIdRef.current = fallbackLineage.turnId;
            hybridInteractionEngine.setCurrentTurnId(fallbackLineage.turnId);
            logService.debug('voice', 'Fallback turn started for pre-transcript tool call');
          }
          const observedToolName = getObservedFirstToolName([funcName]);
          if (observedToolName) {
            logVoiceTurnRoutingOutcome('tool_call', observedToolName);
          }
          const activeTurnId = activeTurnIdRef.current || hybridInteractionEngine.getCurrentTurnId();
          if (activeTurnId && mailDiscussionWorkflowHandledTurnRef.current === activeTurnId) {
            return JSON.stringify({
              success: true,
              workflowHandled: true,
              message: 'The mail discussion reply-all workflow was already handled for this turn.'
            });
          }

          const pendingWorkflowPlan = resolveMailDiscussionReplyAllPlan(undefined, {
            allowPending: true
          });
          const latestUserText = getLatestUserTranscriptTextForTurn(activeTurnId);
          const workflowPlan = pendingWorkflowPlan
            || resolveMailDiscussionReplyAllPlan(latestUserText)
            || resolveMailDiscussionReplyAllPlanFromToolCall(funcName, args);
          if (workflowPlan) {
            maybeEmitVoiceToolAck(funcName, args);
            try {
              const assistantText = await runVoiceCompoundWorkflow(workflowPlan, { speakAssistantText: true });
              return JSON.stringify({
                success: true,
                workflowHandled: true,
                message: assistantText
              });
            } catch (error) {
              const message = error instanceof Error ? error.message : 'Mail discussion workflow failed.';
              logService.error('voice', message);
              return JSON.stringify({
                success: false,
                workflowHandled: true,
                error: message
              });
            }
          }

          maybeEmitVoiceToolAck(funcName, args);
          const output = await handleFunctionCall(funcName, args, currentStore, true);
          maybeEmitVoiceImmediateToolResultAck(funcName, args, output);
          return output;
        }
      },
      voice,
      instructions,
      tools
    );
  }, [
    proxyConfig,
    personality,
    visage,
    setConnectionState,
    setRemoteStream,
    setMicGranted,
    setExpression,
    addTranscript,
    setExpressionWithRevert,
    setAssistantPlaybackState,
    maybeEmitVoiceToolAck,
    maybeEmitVoiceToolCompletionAck,
    maybeEmitVoiceImmediateToolResultAck,
    resetVoiceAckCycle,
    beginVoiceTurnRoutingObservation,
    runVoiceCompoundWorkflow,
    updateConversationLanguageForTurn,
    logVoiceTurnRoutingOutcome
  ]);

  const disconnect = React.useCallback(() => {
    realtimeAudioService.disconnect();
    setRemoteStream(undefined);
    useGrimoireStore.getState().setMicStream(undefined);
    setAssistantPlaybackState('idle');
    setMuted(false);
    hybridInteractionEngine.setVoicePathActive(false);
    hybridInteractionEngine.setAsyncToolCompletionHandler(undefined);
    voiceTurnRoutingRef.current = undefined;
    activeTurnIdRef.current = undefined;
    mailDiscussionWorkflowHandledTurnRef.current = undefined;
    setExpression('idle');
    useGrimoireStore.getState().setConnectionState('idle');
    setConversationLanguage(undefined);
    logService.info('voice', 'Voice audio turned off (text-only mode remains active)');
  }, [setAssistantPlaybackState, setConversationLanguage, setExpression, setMuted, setRemoteStream]);

  const reconnect = React.useCallback(async () => {
    disconnect();
    await Promise.resolve();
    await connect();
  }, [connect, disconnect]);

  const toggleMute = React.useCallback(() => {
    const newMuted = !isMuted;
    realtimeAudioService.setMuted(newMuted);
    setMuted(newMuted);
    logService.info('voice', newMuted ? 'Microphone muted' : 'Microphone unmuted');
  }, [isMuted, setMuted]);

  const sendText = React.useCallback((text: string) => {
    updateConversationLanguageForTurn(text);
    const turnId = beginUserTurnWithIngressPolicy(text);
    activeTurnIdRef.current = turnId;
    hybridInteractionEngine.setCurrentTurnId(turnId);
    // Pre-inject visual state context for text input path
    hybridInteractionEngine.onUserMessage(turnId);
    addTranscript(createTranscriptEntry(text, 'user', turnId));
    logService.debug('llm', `Turn created: ${turnId}`);

    if (realtimeAudioService.isConnected()) {
      // ─── Voice path: WebRTC data channel ───────────────────
      resetVoiceAckCycle();
      beginVoiceTurnRoutingObservation(text);
      realtimeAudioService.sendText(text);
      logService.info('voice', `Text sent (voice): "${text.substring(0, 60)}..."`);
    } else if (textChatServiceRef.current && proxyConfig) {
      // ─── Text path: HTTP chat completions ──────────────────
      const currentStore = useGrimoireStore.getState();
      currentStore.setTextChatActive(true);
      setExpression('thinking');

      // Ensure HIE is initialized for text-only sessions
      if (!hybridInteractionEngine.isInitialized()) {
        const setGazeTarget = currentStore.setGazeTarget;
        nanoServiceRef.current = getNanoService(proxyConfig);
        hybridInteractionEngine.initialize(setExpression, {
          setGazeFn: setGazeTarget,
          nanoService: nanoServiceRef.current,
          sendContextMessage: (ctxText: string, trigger: boolean, contextTurnId?: string) => {
            const tcs = textChatServiceRef.current;
            if (!tcs) return;
            if (contextTurnId) {
              activeTurnIdRef.current = contextTurnId;
              hybridInteractionEngine.setCurrentTurnId(contextTurnId);
            }
            const isToolErrorContext = isHieToolErrorPromptMessage(ctxText);
            if (!trigger) {
              tcs.injectContextMessage(ctxText, false).catch(() => { /* best-effort */ });
              return;
            }
            const iStore = useGrimoireStore.getState();
            if (isToolErrorContext && iStore.textChatActive) {
              // While a text tool-loop is active, tool errors are already part of the
              // function output; inject silently to avoid a duplicate follow-up response.
              tcs.injectContextMessage(ctxText, false).catch(() => { /* best-effort */ });
              return;
            }
            // Interaction-triggered: full response cycle with streaming UI
            iStore.setTextChatActive(true);
            iStore.setExpression('thinking');
            tcs.injectContextMessage(
              ctxText,
              true,
              createStreamingCallbacks(setExpressionWithRevert, 'Context-triggered response', undefined, {
                resolveTurnId: () => contextTurnId || activeTurnIdRef.current
              })
            ).catch((err: Error) => {
              logService.error('llm', `Context injection rejected: ${err.message}`);
              useGrimoireStore.getState().setTextChatActive(false);
            });
          }
        });
      }

      logService.info('llm', `Text sent (HTTP): "${text.substring(0, 60)}..."`);

      textChatServiceRef.current.send(
        text,
        createStreamingCallbacks(setExpressionWithRevert, 'Text response', undefined, {
          surfaceErrors: true,
          resolveTurnId: () => turnId
        })
      ).catch((err: Error) => {
        const s = useGrimoireStore.getState();
        s.setActivityStatus('');
        s.setTextChatActive(false);
        setExpressionWithRevert('confused', 3000);
        logService.error('llm', `TextChatService.send() rejected: ${err.message}`);
      });
    } else {
      logService.warning('system', 'Cannot send text — no voice connection and no proxy config');
    }
  }, [
    addTranscript,
    proxyConfig,
    setExpression,
    setExpressionWithRevert,
    resetVoiceAckCycle,
    beginVoiceTurnRoutingObservation,
    updateConversationLanguageForTurn
  ]);

  const checkHealth = React.useCallback(async () => {
    if (!proxyConfig) {
      setHealthCheckState(false, Date.now(), 'network');
      return;
    }

    try {
      const result = await fetchBackendHealth(proxyConfig, { allowSessionCache: true });
      setHealthCheckState(result.backendOk, result.checkedAt, result.source);
      logService.info('system', `Health check: ${result.backendOk ? 'OK' : 'FAIL'}`, `source=${result.source}`, result.durationMs);
    } catch (error) {
      setHealthCheckState(false, Date.now(), 'network');
      logService.error('system', 'Health check failed', String(error));
    }
  }, [proxyConfig, setHealthCheckState]);

  return { connect, reconnect, disconnect, toggleMute, sendText, checkHealth };
}
