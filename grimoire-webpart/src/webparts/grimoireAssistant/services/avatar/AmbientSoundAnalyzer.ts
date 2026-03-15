/**
 * AmbientSoundAnalyzer — Client-side DSP that classifies ambient mic input
 * into one of four states: silence, voice, typing, or noise.
 *
 * Used to feed the IdleAnimationController:
 *   voice  → force tier 0 (reformation)
 *   typing → cap at tier 1
 *   noise  → cap at tier 2
 *   silence → no cap (idle tiers progress normally)
 *
 * Follows the same connect/disconnect lifecycle pattern as SpeechMouthAnalyzer.
 */

import { acquireSharedAudioContext, type IAudioContextLease } from './SharedAudioContext';

export type AmbientState = 'silence' | 'voice' | 'typing' | 'noise';

export class AmbientSoundAnalyzer {
  private audioContextLease: IAudioContextLease | undefined;
  private audioContext: AudioContext | undefined;
  private analyser: AnalyserNode | undefined;
  private source: MediaStreamAudioSourceNode | undefined;
  private freqData: Uint8Array | undefined;
  private amplitudeHistory: number[] = [];
  private readonly historySize: number = 30;

  constructor(
    private readonly acquireLease: () => IAudioContextLease = acquireSharedAudioContext
  ) {}

  /**
   * Connect to a mic MediaStream.
   * Safe to call multiple times — disconnects first.
   */
  public connect(micStream: MediaStream, audioContextLease?: IAudioContextLease): void {
    this.disconnect();
    try {
      this.audioContextLease = audioContextLease || this.acquireLease();
      this.audioContext = this.audioContextLease.context;
      this.source = this.audioContext.createMediaStreamSource(micStream);
      this.analyser = this.audioContext.createAnalyser();
      this.analyser.fftSize = 256;
      this.analyser.smoothingTimeConstant = 0.5;
      this.source.connect(this.analyser);
      this.freqData = new Uint8Array(this.analyser.frequencyBinCount);
    } catch {
      this.disconnect();
    }
  }

  /** Disconnect and release audio resources. */
  public disconnect(): void {
    if (this.source) {
      try { this.source.disconnect(); } catch { /* already disconnected */ }
      this.source = undefined;
    }
    this.audioContextLease?.release();
    this.audioContextLease = undefined;
    this.audioContext = undefined;
    this.analyser = undefined;
    this.freqData = undefined;
    this.amplitudeHistory = [];
  }

  /** Whether the analyzer is connected and ready to sample. */
  public isConnected(): boolean {
    return !!this.analyser && !!this.freqData;
  }

  /**
   * Sample current mic state and classify.
   * Call once per frame (60fps).
   */
  public sample(): AmbientState {
    if (!this.analyser || !this.freqData) return 'silence';

    this.analyser.getByteFrequencyData(this.freqData);

    // Calculate overall amplitude (RMS of frequency bins)
    let sum = 0;
    for (let i = 0; i < this.freqData.length; i++) {
      sum += this.freqData[i];
    }
    const amplitude = sum / (this.freqData.length * 255);

    // Track amplitude history for variance calculation
    this.amplitudeHistory.push(amplitude);
    if (this.amplitudeHistory.length > this.historySize) {
      this.amplitudeHistory.shift();
    }

    // Silence threshold
    if (amplitude < 0.02) {
      return 'silence';
    }

    // Band energy analysis (5 bands like SpeechMouthAnalyzer)
    const binCount = this.freqData.length;
    const bandSize = Math.floor(binCount / 5);
    let subBass = 0;
    let low = 0;
    let mid = 0;
    let high = 0;
    let veryHigh = 0;

    for (let i = 0; i < bandSize; i++) subBass += this.freqData[i];
    for (let i = bandSize; i < bandSize * 2; i++) low += this.freqData[i];
    for (let i = bandSize * 2; i < bandSize * 3; i++) mid += this.freqData[i];
    for (let i = bandSize * 3; i < bandSize * 4; i++) high += this.freqData[i];
    for (let i = bandSize * 4; i < binCount; i++) veryHigh += this.freqData[i];

    const total = subBass + low + mid + high + veryHigh || 1;

    const lowRatio = (subBass + low) / total;
    const highRatio = (high + veryHigh) / total;

    // Typing: dominated by high frequencies, short transient bursts
    // Amplitude variance tells us about transient vs sustained
    const variance = this.getAmplitudeVariance();

    if (highRatio > 0.45 && variance > 0.001) {
      return 'typing';
    }

    // Voice: dominated by low-mid frequencies, sustained
    if (lowRatio > 0.4 && amplitude > 0.05) {
      return 'voice';
    }

    // Everything else is general noise
    return 'noise';
  }

  private getAmplitudeVariance(): number {
    if (this.amplitudeHistory.length < 5) return 0;
    const mean = this.amplitudeHistory.reduce((a, b) => a + b, 0) / this.amplitudeHistory.length;
    let variance = 0;
    for (const v of this.amplitudeHistory) {
      variance += (v - mean) * (v - mean);
    }
    return variance / this.amplitudeHistory.length;
  }
}
