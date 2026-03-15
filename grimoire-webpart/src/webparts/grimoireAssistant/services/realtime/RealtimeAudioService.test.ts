import {
  RealtimeAudioService,
  shouldDeferImmediateResponse,
  shouldSuppressRealtimeDisplayToolCall
} from './RealtimeAudioService';

describe('RealtimeAudioService deferred function-output responses', () => {
  it('defers immediate response for async placeholder acknowledgments', () => {
    const output = JSON.stringify({
      success: true,
      message: 'Searching for "animals"... Results will appear shortly.'
    });

    expect(shouldDeferImmediateResponse('search_sharepoint', output)).toBe(true);
  });

  it('does not defer for explicit error outputs', () => {
    const output = JSON.stringify({
      success: false,
      error: 'Connection failed'
    });

    expect(shouldDeferImmediateResponse('search_sharepoint', output)).toBe(false);
  });

  it('defers immediate response for expression-only tools', () => {
    const output = JSON.stringify({
      success: true,
      expression: 'happy'
    });

    expect(shouldDeferImmediateResponse('set_expression', output)).toBe(true);
  });

  it('defers immediate response for successful compose forms', () => {
    const output = JSON.stringify({
      success: true,
      message: 'Form displayed in the action panel.'
    });

    expect(shouldDeferImmediateResponse('show_compose_form', output)).toBe(true);
  });

  it('defers immediate response for workflow-handled function outputs', () => {
    const output = JSON.stringify({
      success: true,
      workflowHandled: true,
      message: 'I opened a reply-all draft with the visible recap in the action panel.'
    });

    expect(shouldDeferImmediateResponse('search_emails', output)).toBe(true);
  });

  it('defers immediate response for compose forms even when output is not parseable JSON', () => {
    expect(shouldDeferImmediateResponse('show_compose_form', 'form-opened')).toBe(true);
  });

  it('waits for async function outputs before sending compose results', async () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      callbacks?: {
        onFunctionCall: (callId: string, funcName: string, args: Record<string, unknown>) => Promise<string>;
      };
      handleFunctionCallDone(msg: { call_id?: string; name?: string; arguments?: string; [key: string]: unknown }): void;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.callbacks = {
      onFunctionCall: jest.fn().mockResolvedValue(JSON.stringify({
        success: true,
        message: 'Form displayed in the action panel.'
      }))
    };

    internals.handleFunctionCallDone({
      call_id: 'call_1',
      name: 'show_compose_form',
      arguments: JSON.stringify({ preset: 'email-compose' })
    });

    await Promise.resolve();
    await Promise.resolve();
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(send).toHaveBeenCalledTimes(1);
    expect(send).toHaveBeenCalledWith(JSON.stringify({
      type: 'conversation.item.create',
      item: {
        type: 'function_call_output',
        call_id: 'call_1',
        output: JSON.stringify({
          success: true,
          message: 'Form displayed in the action panel.'
        })
      }
    }));
  });

  it('does not request a follow-up response for expression-only tool outputs', async () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      callbacks?: {
        onFunctionCall: (callId: string, funcName: string, args: Record<string, unknown>) => string | Promise<string>;
      };
      handleFunctionCallDone(msg: { call_id?: string; name?: string; arguments?: string; [key: string]: unknown }): void;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.callbacks = {
      onFunctionCall: jest.fn().mockReturnValue(JSON.stringify({
        success: true
      }))
    };

    internals.handleFunctionCallDone({
      call_id: 'call_1',
      name: 'set_expression',
      arguments: JSON.stringify({ expression: 'happy' })
    });

    await Promise.resolve();
    await Promise.resolve();
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(send).toHaveBeenCalledTimes(1);
    expect(send).toHaveBeenCalledWith(JSON.stringify({
      type: 'conversation.item.create',
      item: {
        type: 'function_call_output',
        call_id: 'call_1',
        output: JSON.stringify({
          success: true
        })
      }
    }));
  });

  it('continues expression-only responses after response.done when no assistant output was produced', async () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      callbacks?: {
        onStateChange: (state: string) => void;
        onTranscript: (text: string, role: 'user' | 'assistant') => void;
        onError: (error: string) => void;
        onFunctionCall: (callId: string, funcName: string, args: Record<string, unknown>) => string | Promise<string>;
      };
      handleFunctionCallDone(msg: { call_id?: string; name?: string; arguments?: string; [key: string]: unknown }): void;
      onDataChannelMessage(event: MessageEvent): void;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.callbacks = {
      onStateChange: jest.fn(),
      onTranscript: jest.fn(),
      onError: jest.fn(),
      onFunctionCall: jest.fn().mockReturnValue(JSON.stringify({
        success: true
      }))
    };

    internals.onDataChannelMessage({ data: JSON.stringify({ type: 'response.created' }) } as MessageEvent);
    internals.handleFunctionCallDone({
      call_id: 'call_1',
      name: 'set_expression',
      arguments: JSON.stringify({ expression: 'thinking' })
    });

    await Promise.resolve();
    await Promise.resolve();
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(send).toHaveBeenCalledTimes(1);
    expect(send).toHaveBeenNthCalledWith(1, JSON.stringify({
      type: 'conversation.item.create',
      item: {
        type: 'function_call_output',
        call_id: 'call_1',
        output: JSON.stringify({
          success: true
        })
      }
    }));

    internals.onDataChannelMessage({ data: JSON.stringify({ type: 'response.done' }) } as MessageEvent);

    expect(send).toHaveBeenCalledTimes(2);
    expect(send).toHaveBeenNthCalledWith(2, JSON.stringify({ type: 'response.create' }));
  });

  it('does not continue expression-only follow-up when the response already produced audio output', async () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      callbacks?: {
        onStateChange: (state: string) => void;
        onTranscript: (text: string, role: 'user' | 'assistant') => void;
        onError: (error: string) => void;
        onFunctionCall: (callId: string, funcName: string, args: Record<string, unknown>) => string | Promise<string>;
      };
      handleFunctionCallDone(msg: { call_id?: string; name?: string; arguments?: string; [key: string]: unknown }): void;
      onDataChannelMessage(event: MessageEvent): void;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.callbacks = {
      onStateChange: jest.fn(),
      onTranscript: jest.fn(),
      onError: jest.fn(),
      onFunctionCall: jest.fn().mockReturnValue(JSON.stringify({
        success: true
      }))
    };

    internals.onDataChannelMessage({ data: JSON.stringify({ type: 'response.created' }) } as MessageEvent);
    internals.onDataChannelMessage({ data: JSON.stringify({ type: 'output_audio_buffer.started' }) } as MessageEvent);
    internals.handleFunctionCallDone({
      call_id: 'call_1',
      name: 'set_expression',
      arguments: JSON.stringify({ expression: 'happy' })
    });

    await Promise.resolve();
    await Promise.resolve();
    await new Promise((resolve) => setTimeout(resolve, 0));

    internals.onDataChannelMessage({ data: JSON.stringify({ type: 'response.done' }) } as MessageEvent);

    expect(send).toHaveBeenCalledTimes(1);
    expect(send).toHaveBeenCalledWith(JSON.stringify({
      type: 'conversation.item.create',
      item: {
        type: 'function_call_output',
        call_id: 'call_1',
        output: JSON.stringify({
          success: true
        })
      }
    }));
  });

  it('queues a deterministic realtime compose ack until the current response completes', () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      responseInFlight: boolean;
      speakDeterministicText(text: string): boolean;
      flushQueuedResponseCreate(): void;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.responseInFlight = true;

    expect(internals.speakDeterministicText('I opened the email draft in the panel.')).toBe(true);
    expect(send).not.toHaveBeenCalled();

    internals.responseInFlight = false;
    internals.flushQueuedResponseCreate();

    expect(send).toHaveBeenCalledWith(JSON.stringify({
      type: 'response.create',
      response: {
        conversation: 'none',
        input: [],
        modalities: ['text', 'audio'],
        tool_choice: 'none',
        max_response_output_tokens: 80,
        instructions: 'Say exactly this sentence and nothing else: "I opened the email draft in the panel."'
      }
    }));
  });
});

describe('RealtimeAudioService cleanup', () => {
  it('clears callbacks before closing RTC resources and remains idempotent', () => {
    const service = new RealtimeAudioService();
    const trackStop = jest.fn();
    const dc = {
      close: jest.fn(),
      onopen: jest.fn(),
      onmessage: jest.fn(),
      onerror: jest.fn(),
      onclose: jest.fn(),
      readyState: 'open'
    } as unknown as RTCDataChannel;
    const pc = {
      close: jest.fn(),
      ontrack: jest.fn(),
      onicecandidate: jest.fn(),
      onconnectionstatechange: jest.fn(),
      oniceconnectionstatechange: jest.fn(),
      ondatachannel: jest.fn()
    } as unknown as RTCPeerConnection;
    const audioElement = {
      srcObject: {},
      onended: jest.fn(),
      onerror: jest.fn(),
      onpause: jest.fn(),
      onplay: jest.fn()
    } as unknown as HTMLAudioElement;
    const mediaStream = {
      getTracks: () => [{ stop: trackStop }]
    } as unknown as MediaStream;
    const internals = service as unknown as {
      cleanup(): void;
      dc?: RTCDataChannel;
      pc?: RTCPeerConnection;
      audioElement?: HTMLAudioElement;
      mediaStream?: MediaStream;
    };

    internals.dc = dc;
    internals.pc = pc;
    internals.audioElement = audioElement;
    internals.mediaStream = mediaStream;

    internals.cleanup();

    expect(trackStop).toHaveBeenCalledTimes(1);
    expect(dc.close).toHaveBeenCalledTimes(1);
    expect(pc.close).toHaveBeenCalledTimes(1);
    expect((dc as { onopen: unknown }).onopen).toBeNull();
    expect((dc as { onmessage: unknown }).onmessage).toBeNull();
    expect((pc as { ontrack: unknown }).ontrack).toBeNull();
    expect((pc as { oniceconnectionstatechange: unknown }).oniceconnectionstatechange).toBeNull();
    expect((audioElement as { onended: unknown }).onended).toBeNull();
    expect((audioElement as { srcObject: unknown }).srcObject).toBeNull();

    expect(() => internals.cleanup()).not.toThrow();
  });
});

describe('RealtimeAudioService playback cutover', () => {
  let originalAudio: typeof Audio | undefined;
  let originalRTCPeerConnection: typeof RTCPeerConnection | undefined;
  let originalMediaDevices: MediaDevices | undefined;
  let originalFetch: typeof fetch | undefined;

  afterEach(() => {
    if (originalAudio) {
      (global as typeof globalThis & { Audio: typeof Audio }).Audio = originalAudio;
    } else {
      Reflect.deleteProperty(global, 'Audio');
    }
    if (originalRTCPeerConnection) {
      (global as typeof globalThis & { RTCPeerConnection: typeof RTCPeerConnection }).RTCPeerConnection = originalRTCPeerConnection;
    } else {
      Reflect.deleteProperty(global, 'RTCPeerConnection');
    }
    if (originalMediaDevices) {
      Object.defineProperty(navigator, 'mediaDevices', {
        configurable: true,
        value: originalMediaDevices
      });
    }
    if (originalFetch) {
      global.fetch = originalFetch;
    }
    jest.restoreAllMocks();
  });

  it('does not mute the realtime remote audio element on connect', async () => {
    const audioElement = {
      autoplay: false,
      muted: true,
      defaultMuted: true,
      playsInline: false,
      srcObject: null,
      play: jest.fn().mockResolvedValue(undefined)
    } as unknown as HTMLAudioElement & {
      defaultMuted?: boolean;
      playsInline?: boolean;
    };
    const dataChannel = {
      readyState: 'open',
      send: jest.fn()
    } as unknown as RTCDataChannel;
    const peerConnection = {
      ontrack: null,
      onconnectionstatechange: null,
      oniceconnectionstatechange: null,
      createDataChannel: jest.fn(() => dataChannel),
      createOffer: jest.fn().mockResolvedValue({ sdp: 'offer-sdp' }),
      setLocalDescription: jest.fn().mockResolvedValue(undefined),
      setRemoteDescription: jest.fn().mockResolvedValue(undefined),
      addTrack: jest.fn()
    } as unknown as RTCPeerConnection;
    const mediaStream = {
      getTracks: jest.fn(() => [{ stop: jest.fn() }]),
      getAudioTracks: jest.fn(() => [])
    } as unknown as MediaStream;

    originalAudio = global.Audio;
    originalRTCPeerConnection = global.RTCPeerConnection;
    originalMediaDevices = navigator.mediaDevices;
    originalFetch = global.fetch;

    (global as typeof globalThis & { Audio: typeof Audio }).Audio = jest.fn(() => audioElement) as unknown as typeof Audio;
    (global as typeof globalThis & { RTCPeerConnection: typeof RTCPeerConnection }).RTCPeerConnection = jest.fn(() => peerConnection) as unknown as typeof RTCPeerConnection;
    Object.defineProperty(navigator, 'mediaDevices', {
      configurable: true,
      value: {
        getUserMedia: jest.fn().mockResolvedValue(mediaStream)
      }
    });
    global.fetch = jest.fn()
      .mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          clientSecret: 'secret',
          expiresAt: '2099-01-01T00:00:00.000Z',
          endpoint: 'https://example.com'
        })
      } as Response)
      .mockResolvedValueOnce({
        ok: true,
        text: async () => 'answer-sdp'
      } as Response);

    const service = new RealtimeAudioService();
    await service.connect(
      {
        proxyUrl: 'https://proxy.example.com',
        proxyApiKey: 'test-key',
        backend: 'reasoning',
        deployment: 'grimoire-realtime',
        apiVersion: '2024-10-21'
      },
      {
        onStateChange: jest.fn(),
        onTranscript: jest.fn(),
        onError: jest.fn(),
        onFunctionCall: jest.fn(() => '{}')
      },
      'alloy',
      'test instructions',
      []
    );

    expect(audioElement.muted).toBe(false);
    expect(audioElement.defaultMuted).toBe(false);
    expect(audioElement.playsInline).toBe(true);
  });

  it('emits assistant playback state changes from realtime output audio events', () => {
    const service = new RealtimeAudioService();
    const playbackStates: string[] = [];
    const internals = service as unknown as {
      callbacks?: {
        onStateChange: (state: string) => void;
        onTranscript: (text: string, role: 'user' | 'assistant') => void;
        onError: (error: string) => void;
        onFunctionCall: (callId: string, funcName: string, args: Record<string, unknown>) => string;
        onAssistantPlaybackStateChange?: (state: 'idle' | 'buffering' | 'playing' | 'error') => void;
      };
      onDataChannelMessage(event: MessageEvent): void;
    };

    internals.callbacks = {
      onStateChange: jest.fn(),
      onTranscript: jest.fn(),
      onError: jest.fn(),
      onFunctionCall: jest.fn(() => '{}'),
      onAssistantPlaybackStateChange: (state) => playbackStates.push(state)
    };

    internals.onDataChannelMessage({ data: JSON.stringify({ type: 'response.created' }) } as MessageEvent);
    internals.onDataChannelMessage({ data: JSON.stringify({ type: 'output_audio_buffer.started' }) } as MessageEvent);
    internals.onDataChannelMessage({ data: JSON.stringify({ type: 'output_audio_buffer.stopped' }) } as MessageEvent);

    expect(playbackStates).toEqual(['buffering', 'playing', 'idle']);
  });

  it('cancels and clears assistant playback on barge-in', () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      responseInFlight: boolean;
      outputAudioBufferActive: boolean;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.responseInFlight = true;
    internals.outputAudioBufferActive = true;

    service.interruptAssistantPlayback();

    expect(send).toHaveBeenNthCalledWith(1, JSON.stringify({ type: 'response.cancel' }));
    expect(send).toHaveBeenNthCalledWith(2, JSON.stringify({ type: 'output_audio_buffer.clear' }));
  });

  it('clears buffered audio without cancelling a completed response', () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      responseInFlight: boolean;
      outputAudioBufferActive: boolean;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.responseInFlight = false;
    internals.outputAudioBufferActive = true;

    service.interruptAssistantPlayback();

    expect(send).toHaveBeenCalledTimes(1);
    expect(send).toHaveBeenCalledWith(JSON.stringify({ type: 'output_audio_buffer.clear' }));
  });
});

describe('RealtimeAudioService barge-in policy', () => {
  it('does not cancel freshly buffered responses before assistant audio starts', () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      responseInFlight: boolean;
      outputAudioBufferActive: boolean;
      responseStartedAt?: number;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.responseInFlight = true;
    internals.outputAudioBufferActive = false;
    internals.responseStartedAt = Date.now();

    service.interruptAssistantPlayback();

    expect(send).not.toHaveBeenCalled();
  });

  it('cancels buffered responses when the user barges in after the grace window', () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      responseInFlight: boolean;
      outputAudioBufferActive: boolean;
      responseStartedAt?: number;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.responseInFlight = true;
    internals.outputAudioBufferActive = false;
    internals.responseStartedAt = Date.now() - 1000;

    service.interruptAssistantPlayback();

    expect(send).toHaveBeenCalledTimes(1);
    expect(send).toHaveBeenCalledWith(JSON.stringify({ type: 'response.cancel' }));
  });
});

describe('RealtimeAudioService watchdogs', () => {
  beforeEach(() => {
    jest.useFakeTimers();
  });

  afterEach(() => {
    jest.useRealTimers();
  });

  it('flushes a queued response when the queue stays stalled after audio has stopped', () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      queuedResponseCreate: boolean;
      responseInFlight: boolean;
      outputAudioBufferActive: boolean;
      scheduleQueuedResponseCreateWatchdog(): void;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;
    internals.queuedResponseCreate = true;
    internals.responseInFlight = false;
    internals.outputAudioBufferActive = false;

    internals.scheduleQueuedResponseCreateWatchdog();
    jest.advanceTimersByTime(2600);

    expect(send).toHaveBeenCalledWith(JSON.stringify({ type: 'response.create' }));
    expect(internals.queuedResponseCreate).toBe(false);
  });

  it('retries response.create when no response.created event arrives', () => {
    const service = new RealtimeAudioService();
    const send = jest.fn();
    const internals = service as unknown as {
      dc?: RTCDataChannel;
      requestResponseCreate(reason: 'user-text' | 'context' | 'function-output' | 'queued'): void;
    };

    internals.dc = {
      readyState: 'open',
      send
    } as unknown as RTCDataChannel;

    internals.requestResponseCreate('function-output');
    jest.advanceTimersByTime(3400);

    expect(send).toHaveBeenCalledTimes(2);
    expect(send).toHaveBeenNthCalledWith(1, JSON.stringify({ type: 'response.create' }));
    expect(send).toHaveBeenNthCalledWith(2, JSON.stringify({ type: 'response.create' }));
  });
});

describe('RealtimeAudioService duplicate display suppression', () => {
  it('suppresses only repeated display tools within the same user turn', () => {
    const displayedTools = new Set<string>(['show_info_card', 'show_markdown']);

    expect(shouldSuppressRealtimeDisplayToolCall('show_info_card', displayedTools)).toBe(true);
    expect(shouldSuppressRealtimeDisplayToolCall('show_markdown', displayedTools)).toBe(true);
    expect(shouldSuppressRealtimeDisplayToolCall('show_selection_list', displayedTools)).toBe(false);
    expect(shouldSuppressRealtimeDisplayToolCall('show_info_card', new Set<string>())).toBe(false);
    expect(shouldSuppressRealtimeDisplayToolCall('search_sharepoint', displayedTools)).toBe(false);
  });
});
