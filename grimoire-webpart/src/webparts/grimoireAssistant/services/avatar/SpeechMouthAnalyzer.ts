/**
 * SpeechMouthAnalyzer
 * Frequency-band speech-mouth cue detection for mouth animation on the Grimm avatar.
 * Uses 5-band energy analysis to classify phoneme-like groups and produce
 * open/width/round mouth parameters with smooth transitions.
 *
 * Speech-mouth cue classification approach and mouth shape definitions inspired by
 * lipsync-engine (MIT License, (c) 2025 Beer Digital LLC)
 * https://github.com/Amoner/lipsync-engine
 *
 * Reimplemented directly on native Web Audio AnalyserNode — zero dependencies,
 * no AudioWorklet required (SPFx-safe).
 */

import { acquireSharedAudioContext, type IAudioContextLease } from './SharedAudioContext';

// ─── Types ─────────────────────────────────────────────────────

export interface IMouthParams {
  /** How open the mouth is (0 = closed, 1 = fully open) */
  openness: number;
  /** Mouth width (0 = narrow, 1 = wide) */
  width: number;
  /** Lip roundness / protrusion (0 = flat, 1 = fully pursed) */
  round: number;
}

/** Internal speech-mouth cues mapped to distinct mouth shapes */
type MouthCue =
  | 'sil'  // Silence
  | 'PP'   // P, B, M — lips pressed
  | 'FF'   // F, V — teeth on lip
  | 'TH'   // TH — tongue between teeth
  | 'DD'   // D, T, N, L — tongue tap
  | 'kk'   // K, G — back tongue
  | 'CH'   // CH, SH, J — rounded fricative
  | 'SS'   // S, Z — sibilant
  | 'nn'   // N, NG — nasal
  | 'RR'   // R — rounded
  | 'aa'   // AA, AH — wide open
  | 'E'    // EH, AE — mid open
  | 'I'    // IH, IY — narrow wide
  | 'O'    // OH, AO — round open
  | 'U';   // UW, OW — pursed

interface IMouthCueShape {
  open: number;
  width: number;
  round: number;
}

interface IScaledMouthCueShape extends IMouthCueShape {
  speechActive: boolean;
}

/** Mouth shapes per cue (adapted from lipsync-engine mouth-shape definitions) */
const MOUTH_CUE_SHAPES: Record<MouthCue, IMouthCueShape> = {
  sil: { open: 0.00, width: 0.50, round: 0.0 },
  PP:  { open: 0.00, width: 0.40, round: 0.0 },
  FF:  { open: 0.05, width: 0.55, round: 0.0 },
  TH:  { open: 0.10, width: 0.50, round: 0.0 },
  DD:  { open: 0.20, width: 0.50, round: 0.0 },
  kk:  { open: 0.25, width: 0.45, round: 0.0 },
  CH:  { open: 0.15, width: 0.35, round: 0.6 },
  SS:  { open: 0.05, width: 0.60, round: 0.0 },
  nn:  { open: 0.15, width: 0.50, round: 0.0 },
  RR:  { open: 0.20, width: 0.40, round: 0.4 },
  aa:  { open: 0.90, width: 0.60, round: 0.0 },
  E:   { open: 0.50, width: 0.65, round: 0.0 },
  I:   { open: 0.25, width: 0.70, round: 0.0 },
  O:   { open: 0.60, width: 0.40, round: 0.8 },
  U:   { open: 0.20, width: 0.30, round: 0.9 }
};

// ─── Frequency Band Extraction ──────────────────────────────────

/** 5-band energy split matching speech formant regions */
interface IBandEnergies {
  sub: number;      // 20-200 Hz — fundamental frequency
  low: number;      // 200-800 Hz — first formant
  mid: number;      // 800-2500 Hz — second formant
  high: number;     // 2500-5500 Hz — fricatives, sibilants
  veryHigh: number; // 5500-12000 Hz — high sibilants
}

/**
 * Map frequency bins to the 5-band energy split.
 * Bins are linearly spaced: bin[i] = i * (sampleRate / fftSize).
 */
function extractBandEnergies(
  data: Uint8Array,
  sampleRate: number,
  fftSize: number
): IBandEnergies {
  const binHz = sampleRate / fftSize;
  const len = data.length;

  // Bin index boundaries for each band
  const bands: Array<[keyof IBandEnergies, number, number]> = [
    ['sub',      20,   200],
    ['low',      200,  800],
    ['mid',      800,  2500],
    ['high',     2500, 5500],
    ['veryHigh', 5500, 12000]
  ];

  const result: IBandEnergies = { sub: 0, low: 0, mid: 0, high: 0, veryHigh: 0 };

  for (const [band, freqLo, freqHi] of bands) {
    const iLo = Math.max(0, Math.floor(freqLo / binHz));
    const iHi = Math.min(len - 1, Math.ceil(freqHi / binHz));
    let sum = 0;
    let count = 0;
    for (let i = iLo; i <= iHi; i++) {
      sum += data[i] / 255;
      count++;
    }
    result[band] = count > 0 ? sum / count : 0;
  }

  return result;
}

// ─── Speech-Mouth Cue Classification ────────────────────────────

/**
 * Classify the current audio frame into a speech-mouth cue based on band energy patterns.
 * Uses formant-like heuristics adapted from lipsync-engine's FrequencyAnalyzer.
 */
function classifyMouthCue(bands: IBandEnergies, amplitude: number): MouthCue {
  const { sub, low, mid, high, veryHigh } = bands;
  const total = sub + low + mid + high + veryHigh;

  if (total < 0.001 || amplitude < 0.01) {
    return 'sil';
  }

  const highRatio = (high + veryHigh) / total;
  const midHighRatio = (mid + high) / total;
  const lowRatio = (sub + low) / total;

  // Sibilants: strong high-frequency energy (S, Z, SH, CH)
  if (highRatio > 0.55) {
    return veryHigh > high ? 'SS' : 'CH';
  }

  // Fricatives: mid + high dominant (F, V, TH)
  if (midHighRatio > 0.50 && high > 0.15) {
    return mid > high ? 'FF' : 'TH';
  }

  // Plosives: high intensity, flat spectrum
  if (amplitude > 0.6) {
    const flatness = 1 - Math.abs(lowRatio - highRatio);
    if (flatness > 0.6) {
      return low > mid ? 'PP' : 'DD';
    }
  }

  // Nasals: strong sub/low, weak mid/high
  if (lowRatio > 0.65 && mid < 0.15) {
    return sub > low ? 'nn' : 'kk';
  }

  // Vowels: classified by formant-like band ratios
  if (sub > low && sub > mid) {
    // Strong fundamental — open vowel
    return 'aa';
  }
  if (low > mid && low > high) {
    // Strong first formant
    return mid > sub ? 'E' : 'O';
  }
  if (mid > low && mid > high) {
    // Strong second formant
    return high > low ? 'I' : 'E';
  }
  if (low > 0.1 && mid > 0.1) {
    return 'RR';
  }

  // Round vowel fallback
  if (low > high) {
    return 'U';
  }

  // Default: generic open mouth
  return 'DD';
}

export function computeTimeDomainAmplitude(data: Uint8Array): number {
  if (data.length === 0) return 0;

  let sumSq = 0;
  for (let i = 0; i < data.length; i++) {
    const centered = (data[i] - 128) / 128;
    sumSq += centered * centered;
  }

  return Math.sqrt(sumSq / data.length);
}

function isOpenVowelCue(cue: MouthCue): boolean {
  return cue === 'aa' || cue === 'E' || cue === 'I' || cue === 'O';
}

function isRoundedCue(cue: MouthCue): boolean {
  return cue === 'O' || cue === 'U' || cue === 'RR' || cue === 'CH';
}

export function resolveSpeechGateThreshold(
  adaptiveNoiseFloor: number,
  silenceThreshold: number
): number {
  return Math.min(0.04, Math.max(silenceThreshold, (adaptiveNoiseFloor * 1.9) + 0.006));
}

function scaleMouthCueShape(
  cue: MouthCue,
  shape: IMouthCueShape,
  signalLevel: number,
  gateThreshold: number
): IScaledMouthCueShape {
  const normalized = Math.max(0, signalLevel - gateThreshold) / Math.max(0.001, 0.12 - gateThreshold);
  const intensity = Math.max(0, Math.min(1, normalized));
  const speechActive = signalLevel >= gateThreshold;
  const openBias = isOpenVowelCue(cue) ? 0.2 : 0.08;
  const widthBias = isOpenVowelCue(cue) ? 0.12 : 0.06;
  const roundBias = isRoundedCue(cue) ? 0.16 : 0.04;
  const intensityFloor = speechActive ? (isOpenVowelCue(cue) ? 0.28 : 0.16) : 0;
  const shapedIntensity = Math.max(intensityFloor, Math.pow(intensity, isOpenVowelCue(cue) ? 0.86 : 1.05));

  return {
    open: Math.min(1, shape.open * (0.52 + (shapedIntensity * 0.88) + openBias)),
    width: 0.5 + ((shape.width - 0.5) * (0.45 + (shapedIntensity * 0.75) + widthBias)),
    round: Math.min(1, shape.round * (0.5 + (shapedIntensity * 0.72) + roundBias)),
    speechActive
  };
}

// ─── Smooth Value Helper ────────────────────────────────────────

function smoothValue(current: number, target: number, factor: number): number {
  return current + (target - current) * factor;
}

// ─── Analyzer ──────────────────────────────────────────────────

export class SpeechMouthAnalyzer {
  private audioContextLease: IAudioContextLease | undefined;
  private audioContext: AudioContext | undefined;
  private analyser: AnalyserNode | undefined;
  private source: MediaStreamAudioSourceNode | undefined;
  private freqData: Uint8Array | undefined;
  private timeData: Uint8Array | undefined;
  private sampleRate: number = 48000;
  private adaptiveNoiseFloor: number = 0.006;
  private speechHangFrames: number = 0;

  /** Smoothed output params */
  private smoothedOpen: number = 0;
  private smoothedWidth: number = 0.5;
  private smoothedRound: number = 0;

  /** Previous cue for hold logic */
  private currentCue: MouthCue = 'sil';
  private holdCounter: number = 0;

  /** Smoothing factor for mouth shape transitions (0-1, higher = faster) */
  private readonly openingSmoothingFactor: number = 0.48;
  private readonly closingSmoothingFactor: number = 0.24;
  private readonly detailSmoothingFactor: number = 0.32;
  /** Below this amplitude, mouth is closed */
  private readonly silenceThreshold: number = 0.015;
  /** Minimum frames before the active cue can switch (prevents flickering) */
  private readonly holdFrames: number = 2;
  /** Short hang to keep the mouth alive between syllables. */
  private readonly maxSpeechHangFrames: number = 4;

  constructor(
    private readonly acquireLease: () => IAudioContextLease = acquireSharedAudioContext
  ) {}

  /**
   * Connect to a MediaStream (typically the WebRTC remote audio stream).
   */
  public connect(stream: MediaStream, audioContextLease?: IAudioContextLease): void {
    this.disconnect();

    try {
      this.audioContextLease = audioContextLease || this.acquireLease();
      this.audioContext = this.audioContextLease.context;
      this.sampleRate = this.audioContext.sampleRate;
      this.source = this.audioContext.createMediaStreamSource(stream);
      this.analyser = this.audioContext.createAnalyser();
      this.analyser.fftSize = 512;
      this.analyser.smoothingTimeConstant = 0.36;
      this.source.connect(this.analyser);
      this.freqData = new Uint8Array(this.analyser.frequencyBinCount);
      this.timeData = new Uint8Array(this.analyser.fftSize);
    } catch {
      this.disconnect();
    }
  }

  /**
   * Disconnect and release all audio resources.
   */
  public disconnect(): void {
    if (this.source) {
      this.source.disconnect();
      this.source = undefined;
    }
    this.audioContextLease?.release();
    this.audioContextLease = undefined;
    this.audioContext = undefined;
    this.analyser = undefined;
    this.freqData = undefined;
    this.timeData = undefined;
    this.adaptiveNoiseFloor = 0.006;
    this.speechHangFrames = 0;
    this.smoothedOpen = 0;
    this.smoothedWidth = 0.5;
    this.smoothedRound = 0;
    this.currentCue = 'sil';
    this.holdCounter = 0;
  }

  /**
   * Sample current audio and return mouth parameters.
   * Call once per frame for smooth animation.
   */
  public sample(): IMouthParams {
    if (!this.analyser || !this.freqData || !this.timeData) {
      // Decay toward closed/neutral
      this.smoothedOpen = smoothValue(this.smoothedOpen, 0, 0.1);
      this.smoothedWidth = smoothValue(this.smoothedWidth, 0.5, 0.1);
      this.smoothedRound = smoothValue(this.smoothedRound, 0, 0.1);
      return { openness: this.smoothedOpen, width: this.smoothedWidth, round: this.smoothedRound };
    }

    // Get frequency data
    this.analyser.getByteFrequencyData(this.freqData);
    this.analyser.getByteTimeDomainData(this.timeData);

    const len = this.freqData.length;
    if (len === 0) {
      return { openness: 0, width: 0.5, round: 0 };
    }

    let sum = 0;
    for (let i = 0; i < len; i++) {
      sum += this.freqData[i] / 255;
    }
    const spectralEnergy = sum / len;
    const amplitude = computeTimeDomainAmplitude(this.timeData);
    const signalLevel = Math.max(amplitude, spectralEnergy * 0.18);
    const nearSilence = signalLevel < (this.adaptiveNoiseFloor * 1.45);
    const noiseBlend = nearSilence ? 0.12 : 0.015;
    this.adaptiveNoiseFloor = smoothValue(this.adaptiveNoiseFloor, signalLevel, noiseBlend);
    const gateThreshold = resolveSpeechGateThreshold(this.adaptiveNoiseFloor, this.silenceThreshold);

    // Silence gate
    if (signalLevel < gateThreshold) {
      if (this.speechHangFrames > 0) {
        this.speechHangFrames--;
      } else {
        const silShape = MOUTH_CUE_SHAPES.sil;
        this.smoothedOpen = smoothValue(this.smoothedOpen, silShape.open, this.closingSmoothingFactor);
        this.smoothedWidth = smoothValue(this.smoothedWidth, silShape.width, this.detailSmoothingFactor);
        this.smoothedRound = smoothValue(this.smoothedRound, silShape.round, this.detailSmoothingFactor);
        this.currentCue = 'sil';
        this.holdCounter = 0;
        return { openness: this.smoothedOpen, width: this.smoothedWidth, round: this.smoothedRound };
      }
    } else {
      this.speechHangFrames = this.maxSpeechHangFrames;
    }

    // Extract 5-band energies
    const bands = extractBandEnergies(this.freqData, this.sampleRate, this.analyser.fftSize);

    // Classify speech-mouth cue
    const detected = classifyMouthCue(bands, signalLevel);

    // Hold logic: prevent rapid flickering between cues
    if (detected !== this.currentCue) {
      this.holdCounter++;
      if (this.holdCounter >= this.holdFrames) {
        this.currentCue = detected;
        this.holdCounter = 0;
      }
    } else {
      this.holdCounter = 0;
    }

    // Get target shape for current cue
    const target = scaleMouthCueShape(this.currentCue, MOUTH_CUE_SHAPES[this.currentCue], signalLevel, gateThreshold);

    const openFactor = target.open > this.smoothedOpen
      ? this.openingSmoothingFactor
      : this.closingSmoothingFactor;
    const detailFactor = target.speechActive
      ? this.detailSmoothingFactor
      : this.closingSmoothingFactor;

    // Smooth toward target
    this.smoothedOpen = smoothValue(this.smoothedOpen, target.open, openFactor);
    this.smoothedWidth = smoothValue(this.smoothedWidth, target.width, detailFactor);
    this.smoothedRound = smoothValue(this.smoothedRound, target.round, detailFactor);

    return {
      openness: this.smoothedOpen,
      width: this.smoothedWidth,
      round: this.smoothedRound
    };
  }

  public isConnected(): boolean {
    return !!this.analyser;
  }
}
