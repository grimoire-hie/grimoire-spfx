export interface IAudioContextLease {
  context: AudioContext;
  release(): void;
}

let sharedAudioContext: AudioContext | undefined;
let activeLeaseCount = 0;

export function acquireSharedAudioContext(): IAudioContextLease {
  if (!sharedAudioContext || sharedAudioContext.state === 'closed') {
    sharedAudioContext = new AudioContext();
  }

  const context = sharedAudioContext;
  activeLeaseCount += 1;
  let released = false;

  return {
    context,
    release(): void {
      if (released) return;
      released = true;
      activeLeaseCount = Math.max(0, activeLeaseCount - 1);

      if (activeLeaseCount === 0 && sharedAudioContext === context && context.state !== 'closed') {
        sharedAudioContext = undefined;
        context.close().catch(() => { /* ignore */ });
      }
    }
  };
}
