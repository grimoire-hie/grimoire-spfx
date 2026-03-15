/**
 * FaceTemplateData
 * Parametric face coordinate generator for the particle avatar.
 * Generates normalized (0-1) coordinate arrays for each FaceRegion.
 */

import { IFaceTemplate } from './ParticleSystem';

export type VisageMode =
  | 'classic'
  | 'anonyMousse'
  | 'robot'
  | 'cat'
  | 'blackCat'
  | 'squirrel'
  | 'particleVisage';

export interface IVisageOption {
  id: VisageMode;
  label: string;
  description: string;
}

export const VISAGE_OPTIONS: Record<VisageMode, IVisageOption> = {
  classic: {
    id: 'classic',
    label: 'GriMoire',
    description: 'Arcane but practical Microsoft 365 steward with calm, lightly mystical delivery'
  },
  anonyMousse: {
    id: 'anonyMousse',
    label: 'AnonyMousse',
    description: 'Guy Fawkes hacker-super-nerd with clever, technical, slightly sly phrasing'
  },
  robot: {
    id: 'robot',
    label: 'Robot',
    description: 'Precise future operator with efficient, quietly advanced language'
  },
  cat: {
    id: 'cat',
    label: 'Majestic Cat',
    description: 'Elegant, polite, slightly distant overseer of Microsoft 365 services'
  },
  blackCat: {
    id: 'blackCat',
    label: 'Black Cat',
    description: 'Panther-like, sleek, elegant, understated assistant presence'
  },
  squirrel: {
    id: 'squirrel',
    label: 'Squirrel',
    description: 'Nimble, funny, lightly mischievous helper that stays competent'
  },
  particleVisage: {
    id: 'particleVisage',
    label: 'Particle Visage',
    description: 'Abstract future human with calm, synthetic, lightly philosophical tone'
  }
};

// ─── Face Parameters ───────────────────────────────────────────

export interface IFaceParams {
  /** Head width multiplier (0.8-1.2) */
  headWidth: number;
  /** Jaw width relative to head (0.6-1.0) */
  jawWidth: number;
  /** Eye spacing from center (0.06-0.12) */
  eyeSpacing: number;
  /** Eye width (0.03-0.06) */
  eyeWidth: number;
  /** Eye height (0.015-0.035) */
  eyeHeight: number;
  /** Eye shape: 0=round, 0.5=almond, 1=narrow */
  eyeShape: number;
  /** Nose width (0.02-0.05) */
  noseWidth: number;
  /** Nose length (0.04-0.08) */
  noseLength: number;
  /** Lip fullness (0.5-1.5) */
  lipFullness: number;
  /** Mouth width (0.04-0.08) */
  mouthWidth: number;
  /** Eyebrow arch (0-1) */
  browArch: number;
  /** Eyebrow thickness (particle count multiplier) */
  browThickness: number;
  /** Face vertical center (0.4-0.5) */
  faceCenter: number;
  /** Chin prominence (0-1) */
  chinProminence: number;
  /** Ear size multiplier */
  earSize: number;
  /** Particle density multiplier (0.5-3.0) */
  density: number;
}

export const DEFAULT_FACE_PARAMS: IFaceParams = {
  headWidth: 1.0,
  jawWidth: 0.8,
  eyeSpacing: 0.14,
  eyeWidth: 0.07,
  eyeHeight: 0.038,
  eyeShape: 0.3,
  noseWidth: 0.045,
  noseLength: 0.09,
  lipFullness: 1.0,
  mouthWidth: 0.105,
  browArch: 0.5,
  browThickness: 1.0,
  faceCenter: 0.43,
  chinProminence: 0.5,
  earSize: 1.0,
  density: 1.0
};

/**
 * Generate subtle random face variation params.
 */
export function randomizeFaceParams(): Partial<IFaceParams> {
  const rand = (min: number, max: number): number => min + Math.random() * (max - min);
  return {
    eyeSpacing: rand(0.12, 0.16),
    eyeShape: rand(0.1, 0.5),
    eyeHeight: rand(0.030, 0.045),
    lipFullness: rand(0.8, 1.2),
    mouthWidth: rand(0.07, 0.11),
    browArch: rand(0.3, 0.7)
  };
}

// ─── Helpers ────────────────────────────────────────────────────

type Point = { x: number; y: number };

function curvePoints(points: Point[], count: number): Point[] {
  const result: Point[] = [];
  const n = Math.max(2, Math.round(count));
  for (let i = 0; i < n; i++) {
    const t = i / (n - 1);
    const segIndex = Math.min(Math.floor(t * (points.length - 1)), points.length - 2);
    const segT = (t * (points.length - 1)) - segIndex;
    const p0 = points[segIndex];
    const p1 = points[segIndex + 1];
    result.push({
      x: p0.x + (p1.x - p0.x) * segT,
      y: p0.y + (p1.y - p0.y) * segT
    });
  }
  return result;
}

function addJitter(points: Point[], amount: number): Point[] {
  return points.map((p) => ({
    x: p.x + (Math.random() - 0.5) * amount,
    y: p.y + (Math.random() - 0.5) * amount
  }));
}

function clamp01(v: number): number {
  if (v < 0) return 0;
  if (v > 1) return 1;
  return v;
}

function clampPoint(pt: Point): Point {
  return { x: clamp01(pt.x), y: clamp01(pt.y) };
}

function centroid(points: Point[]): Point {
  if (points.length === 0) return { x: 0.5, y: 0.5 };
  let sx = 0;
  let sy = 0;
  for (const pt of points) {
    sx += pt.x;
    sy += pt.y;
  }
  return { x: sx / points.length, y: sy / points.length };
}

function indexT(index: number, count: number): number {
  if (count <= 1) return 0;
  return index / (count - 1);
}

function finalizeTemplate(templates: IFaceTemplate[]): IFaceTemplate[] {
  return templates.map((regionTemplate) => ({
    region: regionTemplate.region,
    points: regionTemplate.points.map(clampPoint)
  }));
}

function fillFromBoundary(
  boundary: Point[],
  targetCount: number,
  focusPoints: Point[] = [],
  focusStrength: number = 0.30,
  featureBias: number = 1
): Point[] {
  if (targetCount <= 0 || boundary.length === 0) return [];
  if (targetCount <= boundary.length) {
    return boundary.slice(0, targetCount);
  }

  const center = centroid(boundary);
  const focus = (focusPoints.length > 0 ? focusPoints : [center]).map(clampPoint);

  const points: Point[] = [...boundary];
  const interiorCount = targetCount - boundary.length;
  const boundaryCount = boundary.length;

  for (let i = 0; i < interiorCount; i++) {
    const shell = boundary[i % boundaryCount];
    const seed = Math.sin((i + 1) * 12.9898 + (i % 7) * 78.233) * 43758.5453;
    const hash = seed - Math.floor(seed);
    const radial = 0.10 + Math.pow(hash, 0.58) * 0.86;
    const swirl = Math.sin(i * 0.79) * 0.018;

    let x = center.x + (shell.x - center.x) * (radial + swirl * 0.16);
    let y = center.y + (shell.y - center.y) * (radial - swirl * 0.34);

    const focusPt = focus[i % focus.length];
    const focusSeed = Math.sin((i + 3) * 19.191 + (i % 11) * 47.853) * 12997.481;
    const focusHash = focusSeed - Math.floor(focusSeed);
    const attract = (0.10 + focusHash * 0.28) * focusStrength;
    x += (focusPt.x - x) * attract;
    y += (focusPt.y - y) * attract;

    const eyeBand = Math.exp(-Math.pow((y - (center.y - 0.09)) / 0.07, 2));
    const noseBand = Math.exp(-Math.pow((y - (center.y + 0.01)) / 0.055, 2));
    const mouthBand = Math.exp(-Math.pow((y - (center.y + 0.11)) / 0.08, 2));
    const centerPull = (eyeBand * 0.09 + noseBand * 0.13 + mouthBand * 0.08) * featureBias;
    x += (0.5 - x) * centerPull * 0.36;
    y += (0.53 - y) * noseBand * 0.10 * featureBias;

    x += Math.sin(i * 1.43) * 0.0034;
    y += Math.cos(i * 1.11) * 0.0034;

    points.push(clampPoint({ x, y }));
  }

  return points;
}

function bounds(points: Point[]): { minX: number; maxX: number; minY: number; maxY: number } {
  let minX = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  for (const pt of points) {
    if (pt.x < minX) minX = pt.x;
    if (pt.x > maxX) maxX = pt.x;
    if (pt.y < minY) minY = pt.y;
    if (pt.y > maxY) maxY = pt.y;
  }
  return { minX, maxX, minY, maxY };
}

function createFaceVolumeFocus(boundary: Point[]): Point[] {
  if (boundary.length === 0) return [];
  const b = bounds(boundary);
  const cx = (b.minX + b.maxX) * 0.5;
  const cy = (b.minY + b.maxY) * 0.5;
  const w = Math.max(0.001, b.maxX - b.minX);
  const h = Math.max(0.001, b.maxY - b.minY);
  const cols = [0.18, 0.30, 0.42, 0.50, 0.58, 0.70, 0.82];
  const rows = [0.14, 0.24, 0.34, 0.46, 0.58, 0.70, 0.82];
  const result: Point[] = [];
  for (const row of rows) {
    for (const col of cols) {
      const x = b.minX + w * col;
      const y = b.minY + h * row;
      const nx = (x - cx) / (w * 0.52);
      const ny = (y - cy) / (h * 0.56);
      if ((nx * nx) + (ny * ny) < 1.15) {
        result.push({ x, y });
      }
    }
  }
  result.push({ x: cx, y: b.minY + h * 0.10 });
  result.push({ x: cx, y: b.minY + h * 0.92 });
  return result;
}

function buildHeadContour(boundary: Point[], count: number): Point[] {
  if (boundary.length === 0 || count <= 0) return [];
  const closed = [...boundary, boundary[0]];
  return curvePoints(closed, count);
}

function buildChinCurve(boundary: Point[], count: number): Point[] {
  if (boundary.length === 0 || count <= 0) return [];
  const c = centroid(boundary);
  const lower = boundary
    .filter((pt) => pt.y >= c.y + 0.06)
    .sort((a, b) => a.x - b.x);
  if (lower.length < 3) {
    return curvePoints(
      [
        { x: c.x - 0.11, y: c.y + 0.19 },
        { x: c.x, y: c.y + 0.24 },
        { x: c.x + 0.11, y: c.y + 0.19 }
      ],
      count
    );
  }
  const reduced: Point[] = [];
  const steps = Math.max(6, Math.min(12, lower.length));
  for (let i = 0; i < steps; i++) {
    const t = indexT(i, steps);
    const idx = Math.min(lower.length - 1, Math.floor(t * (lower.length - 1)));
    reduced.push(lower[idx]);
  }
  return curvePoints(reduced, count);
}

function resolveRegionCurve(
  source: Point[] | ((count: number) => Point[]),
  count: number,
  fallback: Point[]
): Point[] {
  const raw = typeof source === 'function' ? source(count) : source;
  if (!raw || raw.length === 0) return fallback;
  if (raw.length === count) return raw;
  if (raw.length < 2) return curvePoints(fallback, count);
  return curvePoints(raw, count);
}

// ─── Shared Topology Counts Across All Visages ────────────────

const EYE_CORE_COUNT = 52;
const EYE_FLOAT_NEAR_COUNT = 26;
const EYE_FLOAT_FAR_COUNT = 20;
const PUPIL_COUNT = 22;
const BROW_COUNT = 28;
const MOUTH_UPPER_MAIN_COUNT = 34;
const MOUTH_LOWER_COUNT = 30;
const MOUTH_FLOAT_COUNT = 34;
const NOSE_COUNT = 28;
const HEAD_CONTOUR_COUNT = 104;
const CHIN_COUNT = 56;
const AMBIENT_COUNT = 1160;

// ─── Classic Template Generator ────────────────────────────────

function generateClassicFaceTemplate(params: Partial<IFaceParams> = {}): IFaceTemplate[] {
  const p = { ...DEFAULT_FACE_PARAMS, ...params };
  const cx = 0.5;
  const cy = p.faceCenter;
  const templates: IFaceTemplate[] = [];

  const eyeY = cy - 0.04;
  const leftEyeCx = cx - p.eyeSpacing;
  const rightEyeCx = cx + p.eyeSpacing;
  const eyeRx = p.eyeWidth * 1.3;
  const eyeRy = p.eyeHeight * (1 - p.eyeShape * 0.3) * 1.2;
  const noseTop = cy + 0.01;
  const mouthY = cy + p.noseLength + 0.045;
  const mouthHalf = p.mouthWidth;
  const lipH = 0.008 * p.lipFullness;

  const eyeCount = EYE_CORE_COUNT;
  const eyeTilt = 0.004;
  const leftEyePts: Point[] = [];
  const rightEyePts: Point[] = [];
  for (let i = 0; i < eyeCount; i++) {
    const angle = (i / eyeCount) * Math.PI * 2;
    const nx = Math.cos(angle);
    const ny = Math.sin(angle);
    const squeeze = 1.0 - 0.34 * Math.pow(Math.abs(nx), 1.5);
    const pinch = 1.0 - 0.24 * Math.pow(Math.abs(nx), 2.0);
    const lidScale = ny < 0 ? 0.84 : 1.07;
    leftEyePts.push({
      x: leftEyeCx + nx * eyeRx,
      y: eyeY + ny * eyeRy * squeeze * pinch * lidScale - nx * eyeTilt
    });
    rightEyePts.push({
      x: rightEyeCx + nx * eyeRx,
      y: eyeY + ny * eyeRy * squeeze * pinch * lidScale + nx * eyeTilt
    });
  }
  templates.push({ region: 'left_eye', points: addJitter(leftEyePts, 0.0026) });
  templates.push({ region: 'right_eye', points: addJitter(rightEyePts, 0.0026) });

  const leftFloatPts: Point[] = [];
  const rightFloatPts: Point[] = [];
  for (let i = 0; i < EYE_FLOAT_NEAR_COUNT; i++) {
    const angle = (i / EYE_FLOAT_NEAR_COUNT) * Math.PI * 2 + Math.random() * 0.4;
    const drift = 1.3 + Math.random() * 0.5;
    leftFloatPts.push({
      x: leftEyeCx + Math.cos(angle) * eyeRx * drift,
      y: eyeY + Math.sin(angle) * eyeRy * drift
    });
    rightFloatPts.push({
      x: rightEyeCx + Math.cos(angle) * eyeRx * drift,
      y: eyeY + Math.sin(angle) * eyeRy * drift
    });
  }
  for (let i = 0; i < EYE_FLOAT_FAR_COUNT; i++) {
    const angle = Math.random() * Math.PI * 2;
    const drift = 2.0 + Math.random() * 1.5;
    leftFloatPts.push({
      x: leftEyeCx + Math.cos(angle) * eyeRx * drift,
      y: eyeY + Math.sin(angle) * eyeRy * drift
    });
    rightFloatPts.push({
      x: rightEyeCx + Math.cos(angle) * eyeRx * drift,
      y: eyeY + Math.sin(angle) * eyeRy * drift
    });
  }
  templates.push({ region: 'left_eye', points: addJitter(leftFloatPts, 0.008) });
  templates.push({ region: 'right_eye', points: addJitter(rightFloatPts, 0.008) });

  const pupilCount = PUPIL_COUNT;
  const pupilR = 0.010;
  const leftPupilPts: Point[] = [];
  const rightPupilPts: Point[] = [];
  for (let i = 0; i < pupilCount; i++) {
    const angle = (i / pupilCount) * Math.PI * 2;
    const r = Math.sqrt(Math.random());
    leftPupilPts.push({
      x: leftEyeCx + Math.cos(angle) * pupilR * r,
      y: eyeY + Math.sin(angle) * pupilR * r
    });
    rightPupilPts.push({
      x: rightEyeCx + Math.cos(angle) * pupilR * r,
      y: eyeY + Math.sin(angle) * pupilR * r
    });
  }
  templates.push({ region: 'left_pupil', points: addJitter(leftPupilPts, 0.002) });
  templates.push({ region: 'right_pupil', points: addJitter(rightPupilPts, 0.002) });

  const browY = eyeY - eyeRy - 0.022;
  const browCount = BROW_COUNT;
  const browRx = p.eyeWidth * 1.5;
  const browArch = 0.014 * (0.7 + p.browArch * 0.6);

  const leftBrowCtrl: Point[] = [
    { x: leftEyeCx - browRx, y: browY + 0.006 },
    { x: leftEyeCx - browRx * 0.35, y: browY - browArch },
    { x: leftEyeCx + browRx * 0.15, y: browY - browArch * 0.7 },
    { x: leftEyeCx + browRx * 0.65, y: browY - browArch * 0.35 },
    { x: leftEyeCx + browRx, y: browY + 0.004 }
  ];
  templates.push({
    region: 'left_eyebrow',
    points: addJitter(curvePoints(leftBrowCtrl, browCount), 0.0025)
  });

  const rightBrowCtrl: Point[] = [
    { x: rightEyeCx - browRx, y: browY + 0.004 },
    { x: rightEyeCx - browRx * 0.65, y: browY - browArch * 0.35 },
    { x: rightEyeCx - browRx * 0.15, y: browY - browArch * 0.7 },
    { x: rightEyeCx + browRx * 0.35, y: browY - browArch },
    { x: rightEyeCx + browRx, y: browY + 0.006 }
  ];
  templates.push({
    region: 'right_eyebrow',
    points: addJitter(curvePoints(rightBrowCtrl, browCount), 0.0025)
  });

  const smileLift = 0.012;
  const cupid = 0.010 * p.lipFullness;
  const upperLipCtrl: Point[] = [
    { x: cx - mouthHalf, y: mouthY - smileLift * 0.75 },
    { x: cx - mouthHalf * 0.62, y: mouthY - smileLift * 0.25 - lipH * 0.65 },
    { x: cx - mouthHalf * 0.28, y: mouthY - lipH * 0.95 },
    { x: cx - mouthHalf * 0.10, y: mouthY - lipH * 0.45 },
    { x: cx, y: mouthY - lipH * 0.75 - cupid },
    { x: cx + mouthHalf * 0.10, y: mouthY - lipH * 0.45 },
    { x: cx + mouthHalf * 0.28, y: mouthY - lipH * 0.95 },
    { x: cx + mouthHalf * 0.62, y: mouthY - smileLift * 0.25 - lipH * 0.65 },
    { x: cx + mouthHalf, y: mouthY - smileLift * 0.75 }
  ];
  templates.push({
    region: 'mouth_upper',
    points: addJitter(curvePoints(upperLipCtrl, MOUTH_UPPER_MAIN_COUNT), 0.0018)
  });

  const lowerLipCtrl: Point[] = [
    { x: cx - mouthHalf, y: mouthY - smileLift * 0.55 },
    { x: cx - mouthHalf * 0.70, y: mouthY - smileLift * 0.20 + lipH * 0.9 },
    { x: cx - mouthHalf * 0.30, y: mouthY + lipH * 1.7 },
    { x: cx, y: mouthY + lipH * 2.1 },
    { x: cx + mouthHalf * 0.30, y: mouthY + lipH * 1.7 },
    { x: cx + mouthHalf * 0.70, y: mouthY - smileLift * 0.20 + lipH * 0.9 },
    { x: cx + mouthHalf, y: mouthY - smileLift * 0.55 }
  ];
  templates.push({
    region: 'mouth_lower',
    points: addJitter(curvePoints(lowerLipCtrl, MOUTH_LOWER_COUNT), 0.0018)
  });

  const mouthFloatPts: Point[] = [];
  const mouthFloatInnerCount = Math.floor(MOUTH_FLOAT_COUNT / 2);
  for (let i = 0; i < mouthFloatInnerCount; i++) {
    const angle = (i / mouthFloatInnerCount) * Math.PI * 2 + Math.random() * 0.3;
    const drift = 1.3 + Math.random() * 0.6;
    mouthFloatPts.push({
      x: cx + Math.cos(angle) * mouthHalf * drift,
      y: mouthY + Math.sin(angle) * 0.02 * drift
    });
  }
  for (let i = mouthFloatInnerCount; i < MOUTH_FLOAT_COUNT; i++) {
    const angle = Math.random() * Math.PI * 2;
    const drift = 2.0 + Math.random() * 1.2;
    mouthFloatPts.push({
      x: cx + Math.cos(angle) * mouthHalf * drift,
      y: mouthY + Math.sin(angle) * 0.03 * drift
    });
  }
  templates.push({
    region: 'mouth_upper',
    points: addJitter(mouthFloatPts, 0.008)
  });

  const noseBridge = p.noseWidth * 0.8;
  const nosePts: Point[] = [
    { x: cx - noseBridge * 0.45, y: noseTop },
    { x: cx + noseBridge * 0.45, y: noseTop + 0.002 },
    { x: cx - noseBridge * 0.28, y: noseTop + p.noseLength * 0.40 },
    { x: cx + noseBridge * 0.28, y: noseTop + p.noseLength * 0.43 },
    { x: cx - noseBridge * 0.90, y: noseTop + p.noseLength * 0.86 },
    { x: cx + noseBridge * 0.90, y: noseTop + p.noseLength * 0.88 }
  ];
  templates.push({
    region: 'nose',
    points: addJitter(curvePoints(nosePts, NOSE_COUNT), 0.0022)
  });

  const ambientBoundaryCount = Math.max(48, Math.floor(AMBIENT_COUNT * 0.34));
  const ambientBoundary: Point[] = [];
  for (let i = 0; i < ambientBoundaryCount; i++) {
    const angle = (i / ambientBoundaryCount) * Math.PI * 2;
    const cheek =
      Math.exp(-Math.pow((angle - 0.72) / 0.42, 2)) * 0.024 +
      Math.exp(-Math.pow((angle - 2.42) / 0.42, 2)) * 0.024;
    const jaw = Math.max(0, Math.sin(angle)) * 0.022;
    const chin = Math.exp(-Math.pow((angle - Math.PI / 2) / 0.22, 2)) * 0.052;
    const forehead = Math.exp(-Math.pow((angle + Math.PI / 2) / 0.52, 2)) * 0.018;
    const r = 0.220 + 0.024 * Math.cos(2 * angle) + cheek + jaw + chin + forehead;
    ambientBoundary.push({
      x: cx + Math.cos(angle) * r,
      y: (cy + 0.040) + Math.sin(angle) * r
    });
  }
  const ambientFocusPoints: Point[] = [
    { x: leftEyeCx, y: eyeY },
    { x: rightEyeCx, y: eyeY },
    { x: cx, y: noseTop + p.noseLength * 0.42 },
    { x: cx, y: mouthY - lipH * 0.6 },
    { x: cx, y: mouthY + lipH * 1.2 },
    { x: cx - mouthHalf * 0.55, y: mouthY },
    { x: cx + mouthHalf * 0.55, y: mouthY }
  ];
  const faceVolumeFocus = createFaceVolumeFocus(ambientBoundary);
  templates.push({
    region: 'head_contour',
    points: addJitter(buildHeadContour(ambientBoundary, HEAD_CONTOUR_COUNT), 0.0018)
  });
  templates.push({
    region: 'chin',
    points: addJitter(buildChinCurve(ambientBoundary, CHIN_COUNT), 0.0018)
  });
  const ambientPts = fillFromBoundary(
    ambientBoundary,
    AMBIENT_COUNT,
    [...ambientFocusPoints, ...faceVolumeFocus],
    0.58,
    1
  );
  templates.push({
    region: 'ambient',
    points: addJitter(ambientPts, 0.0088)
  });

  return finalizeTemplate(templates);
}

// ─── Visage Generators (same topology across all visages) ──────

interface IEyeGeometry {
  cx: number;
  cy: number;
  rx: number;
  ry: number;
  squeeze: number;
  tilt: number;
  cornerPinch: number;
  upperLidScale: number;
  lowerLidScale: number;
  nearDriftMin: number;
  nearDriftMax: number;
  farDriftMin: number;
  farDriftMax: number;
  floatYScale: number;
}

interface IPupilGeometry {
  cx: number;
  cy: number;
  r: number;
}

interface IMouthFloatGeometry {
  cx: number;
  cy: number;
  rx: number;
  ry: number;
  outerRx: number;
  outerRy: number;
}

interface IVisageGeometry {
  leftEye: IEyeGeometry;
  rightEye: IEyeGeometry;
  leftPupil: IPupilGeometry;
  rightPupil: IPupilGeometry;
  leftBrowCtrl: Point[];
  rightBrowCtrl: Point[];
  mouthUpperCtrl: Point[];
  mouthLowerCtrl: Point[];
  mouthFloat: IMouthFloatGeometry;
  nosePoints: Point[];
  ambientPoint: (index: number, count: number) => Point;
  ambientBoundaryRatio?: number;
  ambientFocusPoints?: Point[];
  ambientFocusStrength?: number;
  ambientFeatureBias?: number;
  includeFaceVolumeFocus?: boolean;
  headContourPoints?: Point[] | ((count: number) => Point[]);
  chinPoints?: Point[] | ((count: number) => Point[]);
  jitter: {
    eyeCore: number;
    eyeFloat: number;
    pupil: number;
    brow: number;
    mouthMain: number;
    mouthFloat: number;
    nose: number;
    headContour: number;
    chin: number;
    ambient: number;
  };
}

function buildEyeCorePoints(eye: IEyeGeometry): Point[] {
  const points: Point[] = [];
  const boundaryCount = Math.max(24, Math.floor(EYE_CORE_COUNT * 0.68));
  for (let i = 0; i < boundaryCount; i++) {
    const angle = (i / boundaryCount) * Math.PI * 2;
    const nx = Math.cos(angle);
    const ny = Math.sin(angle);
    const squeeze = 1.0 - eye.squeeze * Math.pow(Math.abs(nx), 1.35);
    const pinch = 1.0 - eye.cornerPinch * Math.pow(Math.abs(nx), 2.0);
    const lidScale = ny < 0 ? eye.upperLidScale : eye.lowerLidScale;
    points.push({
      x: eye.cx + nx * eye.rx,
      y: eye.cy + ny * eye.ry * squeeze * pinch * lidScale + nx * eye.tilt
    });
  }

  const interiorCount = Math.max(0, EYE_CORE_COUNT - boundaryCount);
  for (let i = 0; i < interiorCount; i++) {
    const angle = (i / Math.max(1, interiorCount)) * Math.PI * 2 + (i % 3) * 0.17;
    const nx = Math.cos(angle);
    const ny = Math.sin(angle);
    const seed = Math.sin((i + 1) * 11.713 + (i % 5) * 37.77) * 14358.5453;
    const hash = seed - Math.floor(seed);
    const radial = 0.30 + hash * 0.40;
    const lidScale = ny < 0 ? eye.upperLidScale * 0.92 : eye.lowerLidScale * 1.02;
    points.push({
      x: eye.cx + nx * eye.rx * radial,
      y: eye.cy + ny * eye.ry * radial * lidScale + nx * eye.tilt * 0.45
    });
  }

  return points;
}

function buildEyeFloatPoints(eye: IEyeGeometry): Point[] {
  const points: Point[] = [];
  for (let i = 0; i < EYE_FLOAT_NEAR_COUNT; i++) {
    const angle = (i / EYE_FLOAT_NEAR_COUNT) * Math.PI * 2 + (Math.random() - 0.5) * 0.35;
    const drift = eye.nearDriftMin + Math.random() * (eye.nearDriftMax - eye.nearDriftMin);
    points.push({
      x: eye.cx + Math.cos(angle) * eye.rx * drift,
      y: eye.cy + Math.sin(angle) * eye.ry * drift * eye.floatYScale
    });
  }
  for (let i = 0; i < EYE_FLOAT_FAR_COUNT; i++) {
    const angle = Math.random() * Math.PI * 2;
    const drift = eye.farDriftMin + Math.random() * (eye.farDriftMax - eye.farDriftMin);
    points.push({
      x: eye.cx + Math.cos(angle) * eye.rx * drift,
      y: eye.cy + Math.sin(angle) * eye.ry * drift * eye.floatYScale
    });
  }
  return points;
}

function buildPupilPoints(pupil: IPupilGeometry): Point[] {
  const points: Point[] = [];
  for (let i = 0; i < PUPIL_COUNT; i++) {
    const angle = (i / PUPIL_COUNT) * Math.PI * 2;
    const r = Math.sqrt(Math.random());
    points.push({
      x: pupil.cx + Math.cos(angle) * pupil.r * r,
      y: pupil.cy + Math.sin(angle) * pupil.r * r
    });
  }
  return points;
}

function buildMouthFloatPoints(mouth: IMouthFloatGeometry): Point[] {
  const points: Point[] = [];
  const innerCount = Math.floor(MOUTH_FLOAT_COUNT / 2);
  for (let i = 0; i < innerCount; i++) {
    const angle = (i / innerCount) * Math.PI * 2 + (Math.random() - 0.5) * 0.3;
    points.push({
      x: mouth.cx + Math.cos(angle) * mouth.rx * (1.15 + Math.random() * 0.55),
      y: mouth.cy + Math.sin(angle) * mouth.ry * (1.15 + Math.random() * 0.55)
    });
  }
  for (let i = innerCount; i < MOUTH_FLOAT_COUNT; i++) {
    const angle = Math.random() * Math.PI * 2;
    points.push({
      x: mouth.cx + Math.cos(angle) * mouth.outerRx * (0.95 + Math.random() * 0.60),
      y: mouth.cy + Math.sin(angle) * mouth.outerRy * (0.95 + Math.random() * 0.60)
    });
  }
  return points;
}

function buildVisageTemplate(geometry: IVisageGeometry): IFaceTemplate[] {
  const t: IFaceTemplate[] = [];

  t.push({ region: 'left_eye', points: addJitter(buildEyeCorePoints(geometry.leftEye), geometry.jitter.eyeCore) });
  t.push({ region: 'right_eye', points: addJitter(buildEyeCorePoints(geometry.rightEye), geometry.jitter.eyeCore) });
  t.push({ region: 'left_eye', points: addJitter(buildEyeFloatPoints(geometry.leftEye), geometry.jitter.eyeFloat) });
  t.push({ region: 'right_eye', points: addJitter(buildEyeFloatPoints(geometry.rightEye), geometry.jitter.eyeFloat) });
  t.push({ region: 'left_pupil', points: addJitter(buildPupilPoints(geometry.leftPupil), geometry.jitter.pupil) });
  t.push({ region: 'right_pupil', points: addJitter(buildPupilPoints(geometry.rightPupil), geometry.jitter.pupil) });
  t.push({ region: 'left_eyebrow', points: addJitter(curvePoints(geometry.leftBrowCtrl, BROW_COUNT), geometry.jitter.brow) });
  t.push({ region: 'right_eyebrow', points: addJitter(curvePoints(geometry.rightBrowCtrl, BROW_COUNT), geometry.jitter.brow) });
  t.push({ region: 'mouth_upper', points: addJitter(curvePoints(geometry.mouthUpperCtrl, MOUTH_UPPER_MAIN_COUNT), geometry.jitter.mouthMain) });
  t.push({ region: 'mouth_lower', points: addJitter(curvePoints(geometry.mouthLowerCtrl, MOUTH_LOWER_COUNT), geometry.jitter.mouthMain) });
  t.push({ region: 'mouth_upper', points: addJitter(buildMouthFloatPoints(geometry.mouthFloat), geometry.jitter.mouthFloat) });
  t.push({ region: 'nose', points: addJitter(curvePoints(geometry.nosePoints, NOSE_COUNT), geometry.jitter.nose) });

  const boundaryRatio = Math.max(0.28, Math.min(0.90, geometry.ambientBoundaryRatio ?? 0.34));
  const ambientBoundaryCount = Math.max(48, Math.floor(AMBIENT_COUNT * boundaryRatio));
  const ambientBoundary: Point[] = [];
  for (let i = 0; i < ambientBoundaryCount; i++) {
    ambientBoundary.push(geometry.ambientPoint(i, ambientBoundaryCount));
  }
  const defaultHeadContour = buildHeadContour(ambientBoundary, HEAD_CONTOUR_COUNT);
  const defaultChin = buildChinCurve(ambientBoundary, CHIN_COUNT);
  const headContourPoints = geometry.headContourPoints
    ? resolveRegionCurve(geometry.headContourPoints, HEAD_CONTOUR_COUNT, defaultHeadContour)
    : defaultHeadContour;
  const chinPoints = geometry.chinPoints
    ? resolveRegionCurve(geometry.chinPoints, CHIN_COUNT, defaultChin)
    : defaultChin;
  t.push({
    region: 'head_contour',
    points: addJitter(headContourPoints, geometry.jitter.headContour)
  });
  t.push({
    region: 'chin',
    points: addJitter(chinPoints, geometry.jitter.chin)
  });

  const mouthCenterIdx = Math.floor(geometry.mouthUpperCtrl.length / 2);
  const lowerCenterIdx = Math.floor(geometry.mouthLowerCtrl.length / 2);
  const noseCenterIdx = Math.floor(geometry.nosePoints.length / 2);
  const defaultFocusPoints: Point[] = [
    { x: geometry.leftEye.cx, y: geometry.leftEye.cy },
    { x: geometry.rightEye.cx, y: geometry.rightEye.cy },
    { x: geometry.leftPupil.cx, y: geometry.leftPupil.cy },
    { x: geometry.rightPupil.cx, y: geometry.rightPupil.cy },
    geometry.nosePoints[noseCenterIdx] || { x: 0.5, y: 0.5 },
    geometry.mouthUpperCtrl[mouthCenterIdx] || { x: 0.5, y: 0.62 },
    geometry.mouthLowerCtrl[lowerCenterIdx] || { x: 0.5, y: 0.66 },
    { x: geometry.mouthFloat.cx, y: geometry.mouthFloat.cy }
  ];
  const faceVolumeFocus = createFaceVolumeFocus(ambientBoundary);
  const ambientFocusPoints = geometry.includeFaceVolumeFocus === false
    ? (geometry.ambientFocusPoints ?? defaultFocusPoints)
    : [...(geometry.ambientFocusPoints ?? defaultFocusPoints), ...faceVolumeFocus];
  const ambientPoints = fillFromBoundary(
    ambientBoundary,
    AMBIENT_COUNT,
    ambientFocusPoints,
    geometry.ambientFocusStrength ?? 0.54,
    geometry.ambientFeatureBias ?? 1
  );
  t.push({ region: 'ambient', points: addJitter(ambientPoints, geometry.jitter.ambient) });

  return finalizeTemplate(t);
}

function generateGuyFawkesFaceTemplate(params: Partial<IFaceParams> = {}): IFaceTemplate[] {
  const p = { ...DEFAULT_FACE_PARAMS, ...params };
  const cx = 0.5;
  const eyeSpacing = 0.152 + (p.eyeSpacing - DEFAULT_FACE_PARAMS.eyeSpacing) * 0.24;
  const eyeRx = 0.096 + (p.eyeWidth - DEFAULT_FACE_PARAMS.eyeWidth) * 0.30;
  const eyeRy = 0.016 + (p.eyeHeight - DEFAULT_FACE_PARAMS.eyeHeight) * 0.18;
  const mouthWidth = 0.36 + (p.mouthWidth - DEFAULT_FACE_PARAMS.mouthWidth) * 0.95;
  const noseWidth = 0.014 + (p.noseWidth - DEFAULT_FACE_PARAMS.noseWidth) * 0.20;
  const noseLength = 0.135 + (p.noseLength - DEFAULT_FACE_PARAMS.noseLength) * 0.30;
  const moustacheLift = 0.022;

  return buildVisageTemplate({
    leftEye: {
      cx: cx - eyeSpacing,
      cy: 0.395,
      rx: eyeRx,
      ry: eyeRy,
      squeeze: 0.54,
      tilt: -0.012,
      cornerPinch: 0.60,
      upperLidScale: 0.74,
      lowerLidScale: 1.12,
      nearDriftMin: 1.18,
      nearDriftMax: 1.55,
      farDriftMin: 1.85,
      farDriftMax: 2.70,
      floatYScale: 0.82
    },
    rightEye: {
      cx: cx + eyeSpacing,
      cy: 0.395,
      rx: eyeRx,
      ry: eyeRy,
      squeeze: 0.54,
      tilt: 0.012,
      cornerPinch: 0.60,
      upperLidScale: 0.74,
      lowerLidScale: 1.12,
      nearDriftMin: 1.18,
      nearDriftMax: 1.55,
      farDriftMin: 1.85,
      farDriftMax: 2.70,
      floatYScale: 0.82
    },
    leftPupil: { cx: cx - eyeSpacing, cy: 0.400, r: 0.0078 },
    rightPupil: { cx: cx + eyeSpacing, cy: 0.400, r: 0.0078 },
    leftBrowCtrl: [
      { x: cx - eyeSpacing - 0.118, y: 0.366 },
      { x: cx - eyeSpacing - 0.060, y: 0.300 },
      { x: cx - eyeSpacing + 0.015, y: 0.314 },
      { x: cx - eyeSpacing + 0.075, y: 0.334 },
      { x: cx - eyeSpacing + 0.118, y: 0.368 }
    ],
    rightBrowCtrl: [
      { x: cx + eyeSpacing - 0.118, y: 0.368 },
      { x: cx + eyeSpacing - 0.075, y: 0.334 },
      { x: cx + eyeSpacing - 0.015, y: 0.314 },
      { x: cx + eyeSpacing + 0.060, y: 0.300 },
      { x: cx + eyeSpacing + 0.118, y: 0.366 }
    ],
    mouthUpperCtrl: [
      { x: cx - mouthWidth * 0.5, y: 0.604 },
      { x: cx - mouthWidth * 0.34, y: 0.565 - moustacheLift },
      { x: cx - mouthWidth * 0.22, y: 0.583 },
      { x: cx - mouthWidth * 0.08, y: 0.612 },
      { x: cx, y: 0.626 },
      { x: cx + mouthWidth * 0.08, y: 0.612 },
      { x: cx + mouthWidth * 0.22, y: 0.583 },
      { x: cx + mouthWidth * 0.34, y: 0.565 - moustacheLift },
      { x: cx + mouthWidth * 0.5, y: 0.604 }
    ],
    mouthLowerCtrl: [
      { x: cx - 0.132, y: 0.632 },
      { x: cx - 0.084, y: 0.674 },
      { x: cx - 0.030, y: 0.704 },
      { x: cx, y: 0.718 },
      { x: cx + 0.030, y: 0.704 },
      { x: cx + 0.084, y: 0.674 },
      { x: cx + 0.132, y: 0.632 }
    ],
    mouthFloat: {
      cx,
      cy: 0.648,
      rx: 0.17,
      ry: 0.08,
      outerRx: 0.25,
      outerRy: 0.13
    },
    nosePoints: [
      { x: cx - noseWidth * 0.4, y: 0.438 },
      { x: cx + noseWidth * 0.4, y: 0.440 },
      { x: cx - noseWidth * 0.2, y: 0.438 + noseLength * 0.44 },
      { x: cx + noseWidth * 0.2, y: 0.438 + noseLength * 0.46 },
      { x: cx - noseWidth * 1.30, y: 0.438 + noseLength * 0.94 },
      { x: cx + noseWidth * 1.30, y: 0.438 + noseLength * 0.96 }
    ],
    ambientPoint: (index, count) => {
      const t = indexT(index, count);
      const angle = t * Math.PI * 2;
      const cheekBulge =
        Math.exp(-Math.pow((angle - 0.70) / 0.34, 2)) * 0.045 +
        Math.exp(-Math.pow((angle - 2.44) / 0.34, 2)) * 0.045;
      const chinPoint = Math.exp(-Math.pow((angle - Math.PI / 2) / 0.22, 2)) * 0.094;
      const jawShelf = Math.max(0, Math.sin(angle)) * 0.026;
      const topInset = Math.exp(-Math.pow((angle + Math.PI / 2) / 0.45, 2)) * 0.022;
      const r = 0.214 + 0.014 * Math.cos(2 * angle) + cheekBulge + chinPoint + jawShelf - topInset;
      return {
        x: cx + Math.cos(angle) * r,
        y: 0.472 + Math.sin(angle) * r
      };
    },
    ambientBoundaryRatio: 0.40,
    ambientFocusStrength: 0.62,
    jitter: {
      eyeCore: 0.0022,
      eyeFloat: 0.0065,
      pupil: 0.0015,
      brow: 0.0022,
      mouthMain: 0.0019,
      mouthFloat: 0.0062,
      nose: 0.0018,
      headContour: 0.0019,
      chin: 0.0019,
      ambient: 0.0078
    }
  });
}

export function generateVisageTemplate(
  visage: VisageMode,
  params: Partial<IFaceParams> = {}
): IFaceTemplate[] {
  switch (visage) {
    case 'classic':
      return generateClassicFaceTemplate(params);
    case 'anonyMousse':
      return generateGuyFawkesFaceTemplate(params);
    case 'robot':
      return generateClassicFaceTemplate({
        ...params,
        eyeShape: 1.0,
        eyeWidth: 0.055,
        eyeHeight: 0.022,
        mouthWidth: 0.09,
        browArch: 0.2
      });
    case 'cat':
      return generateClassicFaceTemplate({
        ...params,
        eyeShape: 0.92,
        eyeSpacing: 0.145,
        eyeWidth: 0.06,
        eyeHeight: 0.024,
        mouthWidth: 0.058,
        browArch: 0.16,
        faceCenter: 0.45,
        chinProminence: 0.44
      });
    case 'blackCat':
      return generateClassicFaceTemplate({
        ...params,
        eyeShape: 0.96,
        eyeSpacing: 0.15,
        eyeWidth: 0.062,
        eyeHeight: 0.023,
        mouthWidth: 0.054,
        browArch: 0.12,
        faceCenter: 0.445,
        chinProminence: 0.42
      });
    case 'squirrel':
      return generateClassicFaceTemplate({
        ...params,
        eyeShape: 0.24,
        eyeSpacing: 0.15,
        eyeWidth: 0.066,
        eyeHeight: 0.048,
        mouthWidth: 0.072,
        browArch: 0.52,
        faceCenter: 0.455,
        chinProminence: 0.34
      });
    case 'particleVisage':
      return generateClassicFaceTemplate({
        ...params,
        eyeShape: 1.0,
        eyeSpacing: 0.14,
        eyeWidth: 0.058,
        eyeHeight: 0.018,
        mouthWidth: 0.07,
        browArch: 0.22,
        faceCenter: 0.44,
        chinProminence: 0.58
      });
    default:
      return generateClassicFaceTemplate(params);
  }
}

/**
 * Backward-compatible alias for the original API.
 */
export function generateFaceTemplate(params: Partial<IFaceParams> = {}): IFaceTemplate[] {
  return generateVisageTemplate('classic', params);
}

// ─── Color Themes ───────────────────────────────────────────────

export type ColorTheme = 'blue' | 'red' | 'green' | 'purple' | 'gold' | 'cyan' | 'white';

export interface IColorThemeConfig {
  id: ColorTheme;
  label: string;
  /** Primary particle color */
  primaryColor: string;
  /** Secondary particle color */
  secondaryColor: string;
  /** Glow color */
  glowColor: string;
  /** Background gradient start */
  bgGradientStart: string;
  /** Background gradient end */
  bgGradientEnd: string;
}

export const COLOR_THEMES: Record<ColorTheme, IColorThemeConfig> = {
  blue: {
    id: 'blue',
    label: 'Ice Blue',
    primaryColor: '#a8d8ff',
    secondaryColor: '#6bb8ff',
    glowColor: 'rgba(100, 180, 255, 0.5)',
    bgGradientStart: '#0a0a1a',
    bgGradientEnd: '#0d1525'
  },
  red: {
    id: 'red',
    label: 'Ember',
    primaryColor: '#ff6b6b',
    secondaryColor: '#ff3030',
    glowColor: 'rgba(255, 80, 80, 0.5)',
    bgGradientStart: '#1a0808',
    bgGradientEnd: '#250d0d'
  },
  green: {
    id: 'green',
    label: 'Emerald',
    primaryColor: '#6bffa8',
    secondaryColor: '#30ff80',
    glowColor: 'rgba(80, 255, 130, 0.5)',
    bgGradientStart: '#081a0d',
    bgGradientEnd: '#0d2512'
  },
  purple: {
    id: 'purple',
    label: 'Arcane',
    primaryColor: '#c084fc',
    secondaryColor: '#a855f7',
    glowColor: 'rgba(168, 85, 247, 0.5)',
    bgGradientStart: '#120a1a',
    bgGradientEnd: '#1a0d25'
  },
  gold: {
    id: 'gold',
    label: 'Solar',
    primaryColor: '#ffd700',
    secondaryColor: '#ffaa00',
    glowColor: 'rgba(255, 200, 50, 0.5)',
    bgGradientStart: '#1a1400',
    bgGradientEnd: '#251d05'
  },
  cyan: {
    id: 'cyan',
    label: 'Neon',
    primaryColor: '#00ffff',
    secondaryColor: '#00ccdd',
    glowColor: 'rgba(0, 255, 255, 0.5)',
    bgGradientStart: '#0a1a1a',
    bgGradientEnd: '#0d2525'
  },
  white: {
    id: 'white',
    label: 'Ghost',
    primaryColor: '#e8e8f0',
    secondaryColor: '#c0c0d0',
    glowColor: 'rgba(220, 220, 240, 0.4)',
    bgGradientStart: '#0a0a0e',
    bgGradientEnd: '#121218'
  }
};

/**
 * Get the default face template (single parametric face used for all themes).
 */
export function getDefaultFaceTemplate(): IFaceTemplate[] {
  return generateVisageTemplate('classic');
}
