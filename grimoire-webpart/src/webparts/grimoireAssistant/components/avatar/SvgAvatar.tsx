/**
 * SvgAvatar
 * Pure React + SVG avatar renderer (no Pixi, no Rive).
 */

import * as React from 'react';
import { Expression } from '../../services/avatar/ExpressionEngine';
import { SpeechMouthAnalyzer } from '../../services/avatar/SpeechMouthAnalyzer';
import { COLOR_THEMES, ColorTheme, VisageMode } from '../../services/avatar/FaceTemplateData';
import { PersonalityMode } from '../../services/avatar/PersonalityEngine';
import { IFaceTemplate } from '../../services/avatar/ParticleSystem';
import { AVATAR_MOTION_V2 } from '../../services/avatar/AvatarMotionConfig';
import { AvatarMotionProfile } from '../../services/avatar/AvatarMotionProfile';
import { composeEyeScale, computeMouthTransform } from '../../services/avatar/AvatarMotionMath';
import {
  buildFallbackTwinkleAnchors,
  createTwinkleNodes,
  resolveTwinkleIntensity,
  sampleTwinkleFrame,
  type ITwinkleAnchor,
  type ITwinkleNode
} from '../../services/avatar/AvatarFlourishMotion';
import {
  createSurroundingsNodes,
  isCompactSurroundingsViewport,
  resolveSurroundingsPlacement,
  resolveSurroundingsIntensity,
  sampleSurroundingsNodeFrame,
  type ISurroundingsNode,
  type ISurroundingsPlacement
} from '../../services/avatar/SurroundingsMotion';
import { useGrimoireStore } from '../../store/useGrimoireStore';
import { logService } from '../../services/logging/LogService';
import {
  beginStartupMetric,
  completeStartupMetric
} from '../../services/startup/StartupInstrumentation';
import {
  evaluateAvatarActionCue,
  type IAvatarActionCue,
  type IAvatarCueMouthParams
} from '../../services/avatar/AvatarActionCue';
import classicSvgSource from '../../assets/avatar/svg-source/grimoire_mascot_contract.svg';
import anonyMousseSvgSource from '../../assets/avatar/svg-source/anony_mousse.svg';
import robotSvgSource from '../../assets/avatar/svg-source/robot.svg';
import catSvgSource from '../../assets/avatar/svg-source/cat.svg';
import blackCatSvgSource from '../../assets/avatar/svg-source/black_cat.svg';
import squirrelSvgSource from '../../assets/avatar/svg-source/squirrel.svg';
import particleVisageSvgSource from '../../assets/avatar/svg-source/particle_visage.svg';
import grimoireBookBaseUrl from '../../assets/avatar/png-source/grimoire_book_base.png';
import grimoireBookBrowLeftUrl from '../../assets/avatar/png-source/book_brow_left.png';
import grimoireBookBrowRightUrl from '../../assets/avatar/png-source/book_brow_right.png';
import grimoireBookEyeLeftUrl from '../../assets/avatar/png-source/book_eye_left.png';
import grimoireBookEyeRightUrl from '../../assets/avatar/png-source/book_eye_right.png';
import grimoireBookMouthUrl from '../../assets/avatar/png-source/book_mouth.png';
import robotFaceBaseUrl from '../../assets/avatar/png-source/robot_face_base.png';
import robotBrowLeftUrl from '../../assets/avatar/png-source/robot_brow_left.png';
import robotBrowRightUrl from '../../assets/avatar/png-source/robot_brow_right.png';
import robotEyeLeftUrl from '../../assets/avatar/png-source/robot_eye_left.png';
import robotEyeRightUrl from '../../assets/avatar/png-source/robot_eye_right.png';
import robotMouthUrl from '../../assets/avatar/png-source/robot_mouth.png';
import catBaseUrl from '../../assets/avatar/png-source/cat_base.png';
import catBrowLeftUrl from '../../assets/avatar/png-source/cat_brow_left.png';
import catBrowRightUrl from '../../assets/avatar/png-source/cat_brow_right.png';
import catEyeLeftUrl from '../../assets/avatar/png-source/cat_eye_left.png';
import catEyeRightUrl from '../../assets/avatar/png-source/cat_eye_right.png';
import catMouthUrl from '../../assets/avatar/png-source/cat_mouth.png';
import blackCatBaseUrl from '../../assets/avatar/png-source/black_cat_base.png';
import blackCatBrowLeftUrl from '../../assets/avatar/png-source/black_cat_brow_left.png';
import blackCatBrowRightUrl from '../../assets/avatar/png-source/black_cat_brow_right.png';
import blackCatEyeLeftUrl from '../../assets/avatar/png-source/black_cat_eye_left.png';
import blackCatEyeRightUrl from '../../assets/avatar/png-source/black_cat_eye_right.png';
import blackCatMouthUrl from '../../assets/avatar/png-source/black_cat_mouth.png';
import squirrelBaseUrl from '../../assets/avatar/png-source/squirrel_base.png';
import squirrelBrowLeftUrl from '../../assets/avatar/png-source/squirrel_brow_left.png';
import squirrelBrowRightUrl from '../../assets/avatar/png-source/squirrel_brow_right.png';
import squirrelEyeLeftUrl from '../../assets/avatar/png-source/squirrel_eye_left.png';
import squirrelEyeRightUrl from '../../assets/avatar/png-source/squirrel_eye_right.png';
import squirrelMouthUrl from '../../assets/avatar/png-source/squirrel_mouth.png';
import particleVisageBaseUrl from '../../assets/avatar/png-source/particle_visage_base.png';
import particleVisageBrowLeftUrl from '../../assets/avatar/png-source/particle_visage_brow_left.png';
import particleVisageBrowRightUrl from '../../assets/avatar/png-source/particle_visage_brow_right.png';
import particleVisageEyeLeftUrl from '../../assets/avatar/png-source/particle_visage_eye_left.png';
import particleVisageEyeRightUrl from '../../assets/avatar/png-source/particle_visage_eye_right.png';
import particleVisageMouthUrl from '../../assets/avatar/png-source/particle_visage_mouth_overlay.png';
import { refineBrowMotionForVisage, resolveEyeMotionForVisage } from './SvgAvatarMotionTuning';

const CLASSIC_SVG_SOURCE: string = classicSvgSource;
const ANONY_MOUSSE_SVG_SOURCE: string = anonyMousseSvgSource;
const ROBOT_SVG_SOURCE: string = robotSvgSource;
const CAT_SVG_SOURCE: string = catSvgSource;
const BLACK_CAT_SVG_SOURCE: string = blackCatSvgSource;
const SQUIRREL_SVG_SOURCE: string = squirrelSvgSource;
const PARTICLE_VISAGE_SVG_SOURCE: string = particleVisageSvgSource;

interface IMouthParams {
  openness: number;
  width: number;
  round: number;
}

interface ISvgPart {
  element: SVGGraphicsElement;
  baseTransform: string;
  cx: number;
  cy: number;
}

interface ISvgBindings {
  boundsWidth: number;
  boundsHeight: number;
  mouth?: ISvgPart;
  brows?: ISvgPart;
  eyes?: ISvgPart;
  leftEye?: ISvgPart;
  rightEye?: ISvgPart;
  accentParticles?: ISvgPart;
  faceRoot?: ISvgPart;
  halo?: ISvgPart;
  pageLines?: ISvgPart;
  pages?: ISvgPart;
  pageStack?: ISvgPart;
  covers?: ISvgPart;
}

interface IExpressionModel {
  browOffset: number;
  browRotation: number;
  eyeScale: number;
  mouthLift: number;
  mouthWidthBoost: number;
  mouthOpenBoost: number;
  particleOpacity: number;
}

interface IVisageProfile {
  source: string;
  rootIds: string[];
  accentIds: string[];
  mouthIds: string[];
  browIds: string[];
  leftEyeIds: string[];
  rightEyeIds: string[];
  eyesIds: string[];
}

interface IAvatarTintPalette {
  primary: string;
  secondary: string;
  feature: string;
}

interface IVisageTintProfile {
  primarySelectors: string[];
  secondarySelectors: string[];
  featureSelectors: string[];
}

type SvgReplacementMap = Record<string, string>;

interface ISurroundingsSetup {
  viewBoxWidth: number;
  viewBoxHeight: number;
  placement: ISurroundingsPlacement;
  nodes: ISurroundingsNode[];
  twinkles: ITwinkleNode[];
}

interface IRecentCueTrail {
  type: IAvatarActionCue['type'];
  endedAtMs: number;
}

interface IRectBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

const svgSourceMarkupCache = new Map<string, Promise<string>>();
const imageAssetPreloadCache = new Map<string, Promise<void>>();

const VISAGE_PROFILE: Record<VisageMode, IVisageProfile> = {
  classic: {
    source: CLASSIC_SVG_SOURCE,
    rootIds: ['grimoire_mascot'],
    accentIds: ['halo'],
    mouthIds: ['mouth'],
    browIds: ['brows'],
    leftEyeIds: ['left_eye'],
    rightEyeIds: ['right_eye'],
    eyesIds: []
  },
  anonyMousse: {
    source: ANONY_MOUSSE_SVG_SOURCE,
    rootIds: ['grimoire_mascot'],
    accentIds: ['halo'],
    mouthIds: ['mouth'],
    browIds: ['brows'],
    leftEyeIds: ['left_eye'],
    rightEyeIds: ['right_eye'],
    eyesIds: []
  },
  robot: {
    source: ROBOT_SVG_SOURCE,
    rootIds: ['pixel_ai'],
    accentIds: [],
    mouthIds: ['mouth'],
    browIds: ['brows'],
    leftEyeIds: ['left_eye'],
    rightEyeIds: ['right_eye'],
    eyesIds: []
  },
  cat: {
    source: CAT_SVG_SOURCE,
    rootIds: ['cat_face'],
    accentIds: [],
    mouthIds: ['mouth'],
    browIds: ['brows'],
    leftEyeIds: ['left_eye'],
    rightEyeIds: ['right_eye'],
    eyesIds: []
  },
  blackCat: {
    source: BLACK_CAT_SVG_SOURCE,
    rootIds: ['black_cat_face'],
    accentIds: [],
    mouthIds: ['mouth'],
    browIds: ['brows'],
    leftEyeIds: ['left_eye'],
    rightEyeIds: ['right_eye'],
    eyesIds: []
  },
  squirrel: {
    source: SQUIRREL_SVG_SOURCE,
    rootIds: ['squirrel_face'],
    accentIds: [],
    mouthIds: ['mouth'],
    browIds: ['brows'],
    leftEyeIds: ['left_eye'],
    rightEyeIds: ['right_eye'],
    eyesIds: []
  },
  particleVisage: {
    source: PARTICLE_VISAGE_SVG_SOURCE,
    rootIds: ['particle_face'],
    accentIds: [],
    mouthIds: ['mouth'],
    browIds: ['brows'],
    leftEyeIds: ['left_eye'],
    rightEyeIds: ['right_eye'],
    eyesIds: []
  }
};

const VISAGE_TINT_PROFILE: Record<VisageMode, IVisageTintProfile> = {
  classic: {
    primarySelectors: [],
    secondarySelectors: [],
    featureSelectors: []
  },
  anonyMousse: {
    primarySelectors: [],
    secondarySelectors: [],
    featureSelectors: []
  },
  robot: {
    primarySelectors: [],
    secondarySelectors: [],
    featureSelectors: []
  },
  cat: {
    primarySelectors: [],
    secondarySelectors: [],
    featureSelectors: []
  },
  blackCat: {
    primarySelectors: [],
    secondarySelectors: [],
    featureSelectors: []
  },
  squirrel: {
    primarySelectors: [],
    secondarySelectors: [],
    featureSelectors: []
  },
  particleVisage: {
    primarySelectors: [],
    secondarySelectors: [],
    featureSelectors: []
  }
};

const VISAGE_SVG_REPLACEMENTS: Partial<Record<VisageMode, SvgReplacementMap>> = {
  classic: {
    '__GRIMOIRE_BOOK_BASE__': grimoireBookBaseUrl,
    '__GRIMOIRE_BOOK_BROW_LEFT__': grimoireBookBrowLeftUrl,
    '__GRIMOIRE_BOOK_BROW_RIGHT__': grimoireBookBrowRightUrl,
    '__GRIMOIRE_BOOK_EYE_LEFT__': grimoireBookEyeLeftUrl,
    '__GRIMOIRE_BOOK_EYE_RIGHT__': grimoireBookEyeRightUrl,
    '__GRIMOIRE_BOOK_MOUTH__': grimoireBookMouthUrl
  },
  robot: {
    '__ROBOT_BASE__': robotFaceBaseUrl,
    '__ROBOT_BROW_LEFT__': robotBrowLeftUrl,
    '__ROBOT_BROW_RIGHT__': robotBrowRightUrl,
    '__ROBOT_EYE_LEFT__': robotEyeLeftUrl,
    '__ROBOT_EYE_RIGHT__': robotEyeRightUrl,
    '__ROBOT_MOUTH__': robotMouthUrl
  },
  cat: {
    '__CAT_BASE__': catBaseUrl,
    '__CAT_BROW_LEFT__': catBrowLeftUrl,
    '__CAT_BROW_RIGHT__': catBrowRightUrl,
    '__CAT_EYE_LEFT__': catEyeLeftUrl,
    '__CAT_EYE_RIGHT__': catEyeRightUrl,
    '__CAT_MOUTH__': catMouthUrl
  },
  blackCat: {
    '__BLACK_CAT_BASE__': blackCatBaseUrl,
    '__BLACK_CAT_BROW_LEFT__': blackCatBrowLeftUrl,
    '__BLACK_CAT_BROW_RIGHT__': blackCatBrowRightUrl,
    '__BLACK_CAT_EYE_LEFT__': blackCatEyeLeftUrl,
    '__BLACK_CAT_EYE_RIGHT__': blackCatEyeRightUrl,
    '__BLACK_CAT_MOUTH__': blackCatMouthUrl
  },
  squirrel: {
    '__SQUIRREL_BASE__': squirrelBaseUrl,
    '__SQUIRREL_BROW_LEFT__': squirrelBrowLeftUrl,
    '__SQUIRREL_BROW_RIGHT__': squirrelBrowRightUrl,
    '__SQUIRREL_EYE_LEFT__': squirrelEyeLeftUrl,
    '__SQUIRREL_EYE_RIGHT__': squirrelEyeRightUrl,
    '__SQUIRREL_MOUTH__': squirrelMouthUrl
  },
  particleVisage: {
    '__PARTICLE_VISAGE_BASE__': particleVisageBaseUrl,
    '__PARTICLE_VISAGE_BROW_LEFT__': particleVisageBrowLeftUrl,
    '__PARTICLE_VISAGE_BROW_RIGHT__': particleVisageBrowRightUrl,
    '__PARTICLE_VISAGE_EYE_LEFT__': particleVisageEyeLeftUrl,
    '__PARTICLE_VISAGE_EYE_RIGHT__': particleVisageEyeRightUrl,
    '__PARTICLE_VISAGE_MOUTH__': particleVisageMouthUrl
  }
};

export interface ISvgAvatarProps {
  faceTemplate: IFaceTemplate[];
  visage: VisageMode;
  personality: PersonalityMode;
  expression: Expression;
  remoteStream?: MediaStream;
  micStream?: MediaStream;
  isActive: boolean;
  width?: number;
  height?: number;
  eyeGlow?: { color: string; radius: number; intensity: number };
  colorTheme?: ColorTheme;
  gazeTarget?: 'none' | 'action-panel';
  actionCue?: IAvatarActionCue;
  onActivity?: () => void;
  ambientOnly?: boolean;
}
function resolveMotionIdentity(): string {
  const ctx = useGrimoireStore.getState().userContext;
  const raw = (ctx?.email || ctx?.loginName || '').trim().toLowerCase();
  return raw.length > 0 ? raw : 'anonymous';
}

function applySvgReplacements(svgMarkup: string, visage: VisageMode): string {
  const replacements = VISAGE_SVG_REPLACEMENTS[visage];
  if (!replacements) return svgMarkup;

  return Object.keys(replacements).reduce((markup, token) => (
    markup.split(token).join(replacements[token])
  ), svgMarkup);
}

function resolveReplacementAssetUrls(visage: VisageMode): string[] {
  const replacements = VISAGE_SVG_REPLACEMENTS[visage];
  return replacements ? Object.values(replacements) : [];
}

async function resolveSvgSourceMarkup(source: string): Promise<string> {
  if (source.trim().startsWith('<svg')) return source;

  const cached = svgSourceMarkupCache.get(source);
  if (cached) return cached;

  const pending = fetch(source).then(async (resp) => {
    if (!resp.ok) {
      throw new Error(`Failed to load avatar SVG (${resp.status})`);
    }
    return resp.text();
  });

  svgSourceMarkupCache.set(source, pending);
  try {
    return await pending;
  } catch (error) {
    svgSourceMarkupCache.delete(source);
    throw error;
  }
}

async function preloadImageAsset(url: string): Promise<void> {
  if (!url || typeof Image === 'undefined') return;

  const cached = imageAssetPreloadCache.get(url);
  if (cached) {
    await cached;
    return;
  }

  const pending = new Promise<void>((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve();
    img.onerror = () => reject(new Error(`Failed to preload avatar asset: ${url}`));
    img.src = url;
    if (img.complete) {
      resolve();
    }
  });

  imageAssetPreloadCache.set(url, pending);
  try {
    await pending;
  } catch (error) {
    imageAssetPreloadCache.delete(url);
    throw error;
  }
}

async function preloadVisageAssets(visage: VisageMode): Promise<void> {
  const urls = resolveReplacementAssetUrls(visage);
  if (urls.length === 0) return;
  await Promise.all(urls.map((url) => preloadImageAsset(url)));
}

function clamp01(value: number): number {
  if (value < 0) return 0;
  if (value > 1) return 1;
  return value;
}

function clamp(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function blendMouthParams(base: IMouthParams, overlay: IAvatarCueMouthParams, overlayWeight: number): IMouthParams {
  const weight = clamp01(overlayWeight);
  const baseWeight = 1 - weight;
  return {
    openness: clamp01((base.openness * baseWeight) + (overlay.openness * weight)),
    width: clamp01((base.width * baseWeight) + (overlay.width * weight)),
    round: clamp01((base.round * baseWeight) + (overlay.round * weight))
  };
}

function resolveExpression(expression: Expression): IExpressionModel {
  if (!AVATAR_MOTION_V2.expressionTuningV2) {
    switch (expression) {
      case 'listening':
        return { browOffset: -4, browRotation: 0, eyeScale: 1.05, mouthLift: -1, mouthWidthBoost: 0.03, mouthOpenBoost: 0.06, particleOpacity: 0.55 };
      case 'thinking':
        return { browOffset: 5, browRotation: -2, eyeScale: 0.95, mouthLift: 2, mouthWidthBoost: -0.06, mouthOpenBoost: 0, particleOpacity: 0.38 };
      case 'speaking':
        return { browOffset: -1, browRotation: 0, eyeScale: 1, mouthLift: 0, mouthWidthBoost: 0.08, mouthOpenBoost: 0.22, particleOpacity: 0.7 };
      case 'surprised':
        return { browOffset: -8, browRotation: 0, eyeScale: 1.16, mouthLift: 1, mouthWidthBoost: -0.06, mouthOpenBoost: 0.36, particleOpacity: 0.8 };
      case 'happy':
        return { browOffset: -3, browRotation: 1.5, eyeScale: 0.9, mouthLift: -5, mouthWidthBoost: 0.16, mouthOpenBoost: 0.04, particleOpacity: 0.63 };
      case 'confused':
        return { browOffset: 2, browRotation: -5, eyeScale: 0.96, mouthLift: 2, mouthWidthBoost: -0.12, mouthOpenBoost: 0.03, particleOpacity: 0.46 };
      case 'idle':
      default:
        return { browOffset: 0, browRotation: 0, eyeScale: 1, mouthLift: 0, mouthWidthBoost: 0, mouthOpenBoost: 0, particleOpacity: 0.5 };
    }
  }

  switch (expression) {
    case 'listening':
      return { browOffset: -5, browRotation: 0, eyeScale: 1.08, mouthLift: -1, mouthWidthBoost: 0.04, mouthOpenBoost: 0.07, particleOpacity: 0.58 };
    case 'thinking':
      return { browOffset: 6, browRotation: -3, eyeScale: 0.93, mouthLift: 2, mouthWidthBoost: -0.08, mouthOpenBoost: 0, particleOpacity: 0.4 };
    case 'speaking':
      return { browOffset: -1, browRotation: 0, eyeScale: 0.98, mouthLift: 0, mouthWidthBoost: 0.09, mouthOpenBoost: 0.26, particleOpacity: 0.72 };
    case 'surprised':
      return { browOffset: -9, browRotation: 0, eyeScale: 1.22, mouthLift: 1, mouthWidthBoost: -0.08, mouthOpenBoost: 0.42, particleOpacity: 0.83 };
    case 'happy':
      return { browOffset: -4, browRotation: 1.8, eyeScale: 0.88, mouthLift: -6, mouthWidthBoost: 0.18, mouthOpenBoost: 0.05, particleOpacity: 0.65 };
    case 'confused':
      return { browOffset: 2, browRotation: -6, eyeScale: 0.95, mouthLift: 2, mouthWidthBoost: -0.12, mouthOpenBoost: 0.03, particleOpacity: 0.47 };
    case 'idle':
    default:
      return { browOffset: 0, browRotation: 0, eyeScale: 1, mouthLift: 0, mouthWidthBoost: 0, mouthOpenBoost: 0, particleOpacity: 0.5 };
  }
}

function refineMouthTransformForVisage(visage: VisageMode, transform: ReturnType<typeof computeMouthTransform>): ReturnType<typeof computeMouthTransform> {
  switch (visage) {
    case 'classic':
      return {
        scaleX: clamp(1 + ((transform.scaleX - 1) * 0.24), 0.9, 1.14),
        scaleY: clamp(1 + ((transform.scaleY - 1) * 0.16), 0.94, 1.16),
        liftY: transform.liftY * 0.22,
        jawOffsetY: 0
      };
    case 'robot':
      return {
        scaleX: clamp(1 + ((transform.scaleX - 1) * 0.22), 0.92, 1.12),
        scaleY: clamp(1 + ((transform.scaleY - 1) * 0.14), 0.95, 1.13),
        liftY: transform.liftY * 0.16,
        jawOffsetY: 0
      };
    case 'cat':
      return {
        scaleX: clamp(1 + ((transform.scaleX - 1) * 0.24), 0.95, 1.12),
        scaleY: clamp(1 + ((transform.scaleY - 1) * 0.42), 0.96, 1.38),
        liftY: transform.liftY * 0.22,
        jawOffsetY: 0
      };
    case 'blackCat':
      return {
        scaleX: clamp(1 + ((transform.scaleX - 1) * 0.26), 0.95, 1.14),
        scaleY: clamp(1 + ((transform.scaleY - 1) * 0.56), 0.96, 1.46),
        liftY: transform.liftY * 0.24,
        jawOffsetY: 0
      };
    case 'squirrel':
      return {
        scaleX: clamp(1 + ((transform.scaleX - 1) * 0.22), 0.95, 1.12),
        scaleY: clamp(1 + ((transform.scaleY - 1) * 0.3), 0.96, 1.28),
        liftY: transform.liftY * 0.2,
        jawOffsetY: 0
      };
    case 'particleVisage': {
      const upwardBias = Math.max(0, transform.scaleY - 1) * 48;
      return {
        scaleX: clamp(1 + ((transform.scaleX - 1) * 0.18), 0.97, 1.08),
        scaleY: clamp(1 + ((transform.scaleY - 1) * 0.38), 0.99, 1.22),
        liftY: (transform.liftY * 0.14) - upwardBias,
        jawOffsetY: transform.jawOffsetY * 0.03
      };
    }
    default:
      return transform;
  }
}

function usesOpenMouthOverlay(visage: VisageMode): boolean {
  switch (visage) {
    case 'cat':
    case 'blackCat':
    case 'squirrel':
    case 'particleVisage':
      return true;
    default:
      return false;
  }
}

function resolveMouthOpacityForVisage(
  visage: VisageMode,
  transform: ReturnType<typeof computeMouthTransform>
): number {
  if (!usesOpenMouthOverlay(visage)) return 1;

  switch (visage) {
    case 'cat':
      return clamp01((transform.scaleY - 0.98) / 0.22);
    case 'blackCat':
      return clamp01((transform.scaleY - 0.98) / 0.18);
    case 'squirrel':
      return clamp01((transform.scaleY - 0.98) / 0.2);
    case 'particleVisage':
      return clamp01((transform.scaleY - 1.02) / 0.18);
    default:
      return 1;
  }
}

function bindElement(element: SVGGraphicsElement): ISvgPart {
  const baseTransform = element.getAttribute('transform') || '';
  let cx = 0;
  let cy = 0;
  try {
    const box = element.getBBox();
    cx = box.x + (box.width / 2);
    cy = box.y + (box.height / 2);
  } catch {
    cx = 0;
    cy = 0;
  }
  return { element, baseTransform, cx, cy };
}

function isSvgGraphicsLike(element: Element | undefined): element is SVGGraphicsElement {
  if (!element) return false;
  const candidate = element as unknown as { getBBox?: unknown; setAttribute?: unknown; getAttribute?: unknown };
  return typeof candidate.getBBox === 'function'
    && typeof candidate.setAttribute === 'function'
    && typeof candidate.getAttribute === 'function';
}

function bindPart(root: ParentNode, id: string): ISvgPart | undefined {
  const element = root.querySelector(`[id="${id}"]`) || undefined;
  if (!isSvgGraphicsLike(element)) return undefined;
  return bindElement(element);
}

function bindFirst(root: ParentNode, ids: string[]): ISvgPart | undefined {
  for (const id of ids) {
    const part = bindPart(root, id);
    if (part) return part;
  }
  return undefined;
}

function serializeTransform(part: ISvgPart, tx: number, ty: number, sx: number, sy: number, rotation: number): string {
  const transforms: string[] = [];
  if (part.baseTransform.trim().length > 0) transforms.push(part.baseTransform.trim());
  if (Math.abs(tx) > 0.001 || Math.abs(ty) > 0.001) {
    transforms.push(`translate(${tx.toFixed(3)} ${ty.toFixed(3)})`);
  }
  if (Math.abs(rotation) > 0.001) {
    transforms.push(`rotate(${rotation.toFixed(3)} ${part.cx.toFixed(3)} ${part.cy.toFixed(3)})`);
  }
  if (Math.abs(sx - 1) > 0.001 || Math.abs(sy - 1) > 0.001) {
    transforms.push(`translate(${part.cx.toFixed(3)} ${part.cy.toFixed(3)})`);
    transforms.push(`scale(${sx.toFixed(3)} ${sy.toFixed(3)})`);
    transforms.push(`translate(${(-part.cx).toFixed(3)} ${(-part.cy).toFixed(3)})`);
  }
  return transforms.join(' ');
}

function applyTransform(part: ISvgPart | undefined, tx: number, ty: number, sx: number, sy: number, rotation: number): void {
  if (!part) return;
  if (!part.element.isConnected) return;
  part.element.setAttribute('transform', serializeTransform(part, tx, ty, sx, sy, rotation));
}

function applyOpacity(part: ISvgPart | undefined, opacity: number): void {
  if (!part) return;
  if (!part.element.isConnected) return;
  part.element.style.opacity = `${clamp01(opacity)}`;
}

function extractPartBounds(part: ISvgPart | undefined): IRectBounds | undefined {
  if (!part) return undefined;
  if (!part.element.isConnected) return undefined;
  try {
    const box = part.element.getBBox();
    if (!Number.isFinite(box.x) || !Number.isFinite(box.y) || !Number.isFinite(box.width) || !Number.isFinite(box.height)) {
      return undefined;
    }
    if (box.width <= 0 || box.height <= 0) return undefined;
    return {
      x: box.x,
      y: box.y,
      width: box.width,
      height: box.height
    };
  } catch {
    return undefined;
  }
}

function mergeBounds(parts: Array<IRectBounds | undefined>): IRectBounds | undefined {
  const valid = parts.filter((part): part is IRectBounds => !!part);
  if (valid.length === 0) return undefined;

  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  for (const part of valid) {
    minX = Math.min(minX, part.x);
    minY = Math.min(minY, part.y);
    maxX = Math.max(maxX, part.x + part.width);
    maxY = Math.max(maxY, part.y + part.height);
  }

  return {
    x: minX,
    y: minY,
    width: Math.max(1, maxX - minX),
    height: Math.max(1, maxY - minY)
  };
}

function extractSvgBounds(root: SVGSVGElement): { width: number; height: number } {
  const viewBox = root.getAttribute('viewBox');
  if (!viewBox) return { width: 512, height: 512 };
  const parts = viewBox.trim().split(/\s+/).map((part) => Number(part));
  if (parts.length !== 4 || !Number.isFinite(parts[2]) || !Number.isFinite(parts[3])) {
    return { width: 512, height: 512 };
  }
  return { width: Math.max(1, parts[2]), height: Math.max(1, parts[3]) };
}

function resolveAvatarTintPalette(colorTheme: ColorTheme | undefined): IAvatarTintPalette {
  const theme = colorTheme ? COLOR_THEMES[colorTheme] : COLOR_THEMES.blue;
  return {
    primary: theme.primaryColor,
    secondary: theme.secondaryColor,
    feature: theme.secondaryColor
  };
}

function recolorSvgElement(element: Element, color: string): void {
  const targetTags = 'path,circle,ellipse,polygon,rect,line,polyline';
  const candidates: Element[] = [];
  if (element.matches(targetTags)) {
    candidates.push(element);
  }
  candidates.push(...Array.from(element.querySelectorAll(targetTags)));

  for (const candidate of candidates) {
    const node = candidate as unknown as {
      getAttribute?: (name: string) => string | null;
      setAttribute?: (name: string, value: string) => void;
      tagName?: string;
    };
    if (typeof node.getAttribute !== 'function' || typeof node.setAttribute !== 'function' || typeof node.tagName !== 'string') {
      continue;
    }
    const fill = node.getAttribute('fill');
    const stroke = node.getAttribute('stroke');
    const tag = node.tagName.toLowerCase();

    if (fill && fill.toLowerCase() !== 'none') {
      node.setAttribute('fill', color);
    }
    if (stroke && stroke.toLowerCase() !== 'none') {
      node.setAttribute('stroke', color);
    }

    if (!fill && !stroke) {
      if (tag === 'line' || tag === 'path' || tag === 'polyline') {
        node.setAttribute('stroke', color);
      } else {
        node.setAttribute('fill', color);
      }
    }
  }
}

function recolorSelectors(root: ParentNode, selectors: string[], color: string): void {
  for (const selector of selectors) {
    const targets = Array.from(root.querySelectorAll(selector));
    for (const target of targets) {
      recolorSvgElement(target, color);
    }
  }
}

function applyAvatarTint(root: ParentNode, visage: VisageMode, palette: IAvatarTintPalette): void {
  const tintProfile = VISAGE_TINT_PROFILE[visage] ?? VISAGE_TINT_PROFILE.classic;
  recolorSelectors(root, tintProfile.primarySelectors, palette.primary);
  recolorSelectors(root, tintProfile.secondarySelectors, palette.secondary);
  recolorSelectors(root, tintProfile.featureSelectors, palette.feature);
}

function applySvgElementOpacity(element: SVGElement | undefined, opacity: number): void {
  if (!element) return;
  if (!element.isConnected) return;
  element.style.opacity = `${clamp01(opacity)}`;
}

function applySvgElementsOpacity(elements: Array<SVGElement | undefined>, opacity: number): void {
  for (const element of elements) {
    applySvgElementOpacity(element, opacity);
  }
}

function resolvePaletteColor(colorRole: 'primary' | 'secondary' | 'feature', palette: IAvatarTintPalette): string {
  switch (colorRole) {
    case 'secondary':
      return palette.secondary;
    case 'feature':
      return palette.feature;
    case 'primary':
    default:
      return palette.primary;
  }
}

function pushTwinkleAnchor(
  anchors: ITwinkleAnchor[],
  viewport: { width: number; height: number },
  x: number,
  y: number,
  swayX: number,
  swayY: number,
  colorRole: 'primary' | 'secondary' | 'feature',
  cueAffinity: number
): void {
  if (!Number.isFinite(x) || !Number.isFinite(y)) return;
  anchors.push({
    x: clamp(x, 12, viewport.width - 12),
    y: clamp(y, 12, viewport.height - 12),
    swayX,
    swayY,
    colorRole,
    cueAffinity
  });
}

function buildTwinkleAnchors(
  visage: VisageMode,
  bounds: IRectBounds,
  viewport: { width: number; height: number },
  haloBounds: IRectBounds | undefined,
  pageLinesBounds: IRectBounds | undefined,
  pageStackBounds: IRectBounds | undefined,
  pagesBounds: IRectBounds | undefined
): ITwinkleAnchor[] {
  const anchors: ITwinkleAnchor[] = [];
  if (visage !== 'classic') {
    return buildFallbackTwinkleAnchors(bounds, viewport);
  }

  if (haloBounds) {
    pushTwinkleAnchor(
      anchors,
      viewport,
      haloBounds.x + (haloBounds.width * 0.5),
      haloBounds.y + (haloBounds.height * 0.08),
      2.2,
      3.6,
      'feature',
      1
    );
    pushTwinkleAnchor(
      anchors,
      viewport,
      haloBounds.x + (haloBounds.width * 0.12),
      haloBounds.y + (haloBounds.height * 0.56),
      2.8,
      2.4,
      'secondary',
      0.82
    );
    pushTwinkleAnchor(
      anchors,
      viewport,
      haloBounds.x + (haloBounds.width * 0.88),
      haloBounds.y + (haloBounds.height * 0.56),
      2.8,
      2.4,
      'secondary',
      0.82
    );
  }

  if (pageLinesBounds) {
    pushTwinkleAnchor(
      anchors,
      viewport,
      pageLinesBounds.x + (pageLinesBounds.width * 0.08),
      pageLinesBounds.y + (pageLinesBounds.height * 0.14),
      2.5,
      1.8,
      'primary',
      0.68
    );
    pushTwinkleAnchor(
      anchors,
      viewport,
      pageLinesBounds.x + (pageLinesBounds.width * 0.92),
      pageLinesBounds.y + (pageLinesBounds.height * 0.16),
      2.5,
      1.8,
      'primary',
      0.68
    );
  }

  if (pagesBounds) {
    pushTwinkleAnchor(
      anchors,
      viewport,
      pagesBounds.x + (pagesBounds.width * 0.5),
      pagesBounds.y + (pagesBounds.height * 0.12),
      2.4,
      2,
      'feature',
      0.76
    );
  }

  if (pageStackBounds) {
    pushTwinkleAnchor(
      anchors,
      viewport,
      pageStackBounds.x + (pageStackBounds.width * 0.16),
      pageStackBounds.y + (pageStackBounds.height * 0.2),
      2.2,
      1.6,
      'secondary',
      0.54
    );
    pushTwinkleAnchor(
      anchors,
      viewport,
      pageStackBounds.x + (pageStackBounds.width * 0.84),
      pageStackBounds.y + (pageStackBounds.height * 0.2),
      2.2,
      1.6,
      'secondary',
      0.54
    );
  }

  return [...anchors, ...buildFallbackTwinkleAnchors(bounds, viewport)];
}

function resetClassicParallax(bindings: ISvgBindings | undefined): void {
  if (!bindings) return;
  applyTransform(bindings.halo, 0, 0, 1, 1, 0);
  applyTransform(bindings.pages, 0, 0, 1, 1, 0);
  applyTransform(bindings.covers, 0, 0, 1, 1, 0);
  applyTransform(bindings.pageLines, 0, 0, 1, 1, 0);
  applyTransform(bindings.pageStack, 0, 0, 1, 1, 0);
}

function resolveEffectiveExpression(
  expression: Expression,
  assistantPlaybackState: 'idle' | 'buffering' | 'playing' | 'error'
): Expression {
  if (assistantPlaybackState === 'playing' && expression !== 'confused') {
    return 'speaking';
  }
  return expression;
}

function resolvePresentationTransformStyle(visage: VisageMode): React.CSSProperties {
  // GriMoire and Robot get some effective top margin from their asset geometry,
  // but AnonyMousse has very little intrinsic top padding, so large negative
  // offsets become visible immediately on shorter viewports.
  switch (visage) {
    case 'classic':
      return {
        transform: 'translateY(-2%) scale(0.9)',
        transformOrigin: 'center center'
      };
    case 'anonyMousse':
      return {
        transform: 'translateY(-5%) scale(0.67)',
        transformOrigin: 'center center'
      };
    case 'robot':
      return {
        transform: 'translateY(-2%) scale(0.7)',
        transformOrigin: 'center center'
      };
    case 'cat':
      return {
        transform: 'translateY(-4%) scale(0.62)',
        transformOrigin: 'center center'
      };
    case 'blackCat':
      return {
        transform: 'translateY(-1%) scale(0.72)',
        transformOrigin: 'center center'
      };
    case 'squirrel':
      return {
        transform: 'translateY(-4%) scale(0.63)',
        transformOrigin: 'center center'
      };
    case 'particleVisage':
      return {
        transform: 'translateY(-1%) scale(0.74)',
        transformOrigin: 'center center'
      };
    default:
      return {};
  }
}

function resolveAvatarVisualStyle(visage: VisageMode): React.CSSProperties {
  switch (visage) {
    case 'classic':
      return { opacity: 0.96 };
    case 'anonyMousse':
      return { opacity: 0.72 };
    case 'particleVisage':
      return { opacity: 0.92 };
    default:
      return {};
  }
}

export const SvgAvatar: React.FC<ISvgAvatarProps> = ({
  visage,
  personality: _personality,
  expression,
  remoteStream,
  isActive,
  width,
  height,
  colorTheme,
  gazeTarget,
  actionCue,
  onActivity
}) => {
  const svgHostRef = React.useRef<HTMLDivElement>(null);
  const bindingsRef = React.useRef<ISvgBindings | undefined>(undefined);
  const surroundingsAuraRef = React.useRef<SVGEllipseElement | undefined>(undefined);
  const surroundingsNodeRefs = React.useRef<Array<SVGCircleElement | undefined>>([]);
  const surroundingsTwinkleRefs = React.useRef<Array<SVGGElement | undefined>>([]);
  const frameRef = React.useRef<number>(0);
  const speechMouthRef = React.useRef<SpeechMouthAnalyzer>(new SpeechMouthAnalyzer());
  const latestActionCueRef = React.useRef<IAvatarActionCue | undefined>(actionCue);
  const activeActionCueRef = React.useRef<IAvatarActionCue | undefined>(undefined);
  const activeActionCueIdRef = React.useRef<number | undefined>(undefined);
  const missingBindingsCueIdRef = React.useRef<number | undefined>(undefined);
  const recentCueTrailRef = React.useRef<IRecentCueTrail | undefined>(undefined);
  const lastBindingsSignatureRef = React.useRef<string>('');
  const motionProfileRef = React.useRef<AvatarMotionProfile | undefined>(undefined);
  const firstAnimatedFrameReadyRef = React.useRef<boolean>(false);
  const firstAnimatedMetricPendingRef = React.useRef<boolean>(false);
  const [inlineSvgMarkup, setInlineSvgMarkup] = React.useState<string>('');
  const [renderedVisage, setRenderedVisage] = React.useState<VisageMode>(visage);
  const [surroundingsSetup, setSurroundingsSetup] = React.useState<ISurroundingsSetup | undefined>(undefined);
  const [isCompactViewport, setIsCompactViewport] = React.useState(
    typeof window !== 'undefined' ? isCompactSurroundingsViewport(window.innerWidth) : false
  );
  const [isDocumentHidden, setIsDocumentHidden] = React.useState(
    typeof document !== 'undefined' ? document.hidden : false
  );
  const assistantPlaybackState = useGrimoireStore((s) => s.assistantPlaybackState);
  const setAvatarRenderState = useGrimoireStore((s) => s.setAvatarRenderState);
  const motionIdentity = useGrimoireStore((s) => {
    const raw = (s.userContext?.email || s.userContext?.loginName || '').trim().toLowerCase();
    return raw.length > 0 ? raw : 'anonymous';
  });

  const motionSeedKey = React.useMemo(
    () => `${motionIdentity || resolveMotionIdentity()}|${renderedVisage}`,
    [motionIdentity, renderedVisage]
  );

  const requestedProfile = React.useMemo(() => VISAGE_PROFILE[visage] ?? VISAGE_PROFILE.classic, [visage]);
  const profile = React.useMemo(() => VISAGE_PROFILE[renderedVisage] ?? VISAGE_PROFILE.classic, [renderedVisage]);
  const tintPalette = React.useMemo(
    () => resolveAvatarTintPalette(colorTheme),
    [colorTheme]
  );
  const presentationTransformStyle = React.useMemo<React.CSSProperties>(
    () => resolvePresentationTransformStyle(renderedVisage),
    [renderedVisage]
  );
  const effectiveExpression = React.useMemo<Expression>(
    () => resolveEffectiveExpression(expression, assistantPlaybackState),
    [assistantPlaybackState, expression]
  );
  const avatarVisualStyle = React.useMemo<React.CSSProperties>(
    () => resolveAvatarVisualStyle(renderedVisage),
    [renderedVisage]
  );
  const auraGradientId = React.useMemo(
    () => `grimoire-surroundings-aura-${Math.random().toString(36).slice(2, 10)}`,
    []
  );
  const source = requestedProfile.source;

  React.useEffect(() => {
    if (typeof window === 'undefined') return undefined;

    const handleResize = (): void => {
      const next = isCompactSurroundingsViewport(window.innerWidth);
      setIsCompactViewport((current) => (current === next ? current : next));
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  React.useEffect(() => {
    if (typeof document === 'undefined') return undefined;

    const handleVisibilityChange = (): void => {
      setIsDocumentHidden(document.hidden);
    };

    document.addEventListener('visibilitychange', handleVisibilityChange);
    return () => document.removeEventListener('visibilitychange', handleVisibilityChange);
  }, []);

  React.useEffect(() => {
    let cancelled = false;
    setAvatarRenderState('svg-loading');
    firstAnimatedFrameReadyRef.current = false;
    firstAnimatedMetricPendingRef.current = false;
    setInlineSvgMarkup('');
    beginStartupMetric('svg-fetch', `visage=${visage}`);
    const loadSvg = async (): Promise<void> => {
      try {
        const [svgText] = await Promise.all([
          resolveSvgSourceMarkup(source),
          preloadVisageAssets(visage)
        ]);

        const resolvedMarkup = applySvgReplacements(svgText, visage);
        if (!cancelled) {
          completeStartupMetric('svg-fetch', `visage=${visage}`);
          beginStartupMetric('svg-bind', `visage=${visage}`);
          setRenderedVisage(visage);
          setInlineSvgMarkup(resolvedMarkup);
          logService.debug('system', `Avatar SVG loaded for ${visage}`, `source=${source}, length=${resolvedMarkup.length}`);
        }
      } catch (error) {
        if (!cancelled) {
          const message = error instanceof Error ? error.message : 'unknown';
          completeStartupMetric('svg-fetch', `visage=${visage}; error=${message}`);
          logService.warning('system', `Avatar SVG load failed for ${visage}`, `source=${source}`);
        }
      }
    };

    loadSvg().catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [setAvatarRenderState, source, visage]);

  React.useEffect(() => {
    const host = svgHostRef.current;
    if (!host) {
      bindingsRef.current = undefined;
      setSurroundingsSetup(undefined);
      return;
    }
    const rootSvg = host.querySelector('svg');
    if (!(rootSvg instanceof SVGSVGElement)) {
      bindingsRef.current = undefined;
      setSurroundingsSetup(undefined);
      return;
    }
    rootSvg.setAttribute('width', '100%');
    rootSvg.setAttribute('height', '100%');

    const bounds = extractSvgBounds(rootSvg);
    const faceRoot = bindFirst(rootSvg, profile.rootIds);
    const mouth = bindFirst(rootSvg, profile.mouthIds);
    const brows = bindFirst(rootSvg, profile.browIds);
    const leftEye = bindFirst(rootSvg, profile.leftEyeIds);
    const rightEye = bindFirst(rootSvg, profile.rightEyeIds);
    const eyes = bindFirst(rootSvg, profile.eyesIds);

    bindingsRef.current = {
      boundsWidth: bounds.width,
      boundsHeight: bounds.height,
      mouth,
      brows,
      eyes,
      leftEye,
      rightEye,
      accentParticles: bindFirst(rootSvg, profile.accentIds),
      faceRoot,
      halo: bindPart(rootSvg, 'halo'),
      pageLines: bindPart(rootSvg, 'page_lines'),
      pages: bindPart(rootSvg, 'pages'),
      pageStack: bindPart(rootSvg, 'page_stack'),
      covers: bindPart(rootSvg, 'covers')
    };
    const b = bindingsRef.current;
    const signature = `L${!!b.leftEye}|R${!!b.rightEye}|E${!!b.eyes}|M${!!b.mouth}|B${!!b.brows}|P${!!b.accentParticles}|F${!!b.faceRoot}|H${!!b.halo}|PL${!!b.pageLines}|PG${!!b.pages}|PS${!!b.pageStack}|C${!!b.covers}`;
    if (signature !== lastBindingsSignatureRef.current) {
      lastBindingsSignatureRef.current = signature;
      logService.info('system', `Avatar bindings ready (${renderedVisage})`, signature);
    }
    setAvatarRenderState('bindings-ready');
    completeStartupMetric('svg-bind', `visage=${renderedVisage}`);
    if (!firstAnimatedMetricPendingRef.current) {
      firstAnimatedMetricPendingRef.current = true;
      beginStartupMetric('first-animated-frame', `visage=${renderedVisage}`);
    }
    applyAvatarTint(rootSvg, renderedVisage, tintPalette);
    applyOpacity(bindingsRef.current.mouth, usesOpenMouthOverlay(renderedVisage) ? 0 : 1);

    if (!AVATAR_MOTION_V2.ambientAuraV1 && !AVATAR_MOTION_V2.surroundingsV2 && !AVATAR_MOTION_V2.twinklesV1) {
      setSurroundingsSetup(undefined);
      return;
    }

    const featureBounds = mergeBounds([
      extractPartBounds(brows),
      extractPartBounds(leftEye),
      extractPartBounds(rightEye),
      extractPartBounds(eyes),
      extractPartBounds(mouth)
    ]);
    const surroundingsBounds = featureBounds || extractPartBounds(faceRoot);
    const haloBounds = extractPartBounds(b.halo);
    const pageLinesBounds = extractPartBounds(b.pageLines);
    const pageStackBounds = extractPartBounds(b.pageStack);
    const pagesBounds = extractPartBounds(b.pages);
    if (surroundingsBounds) {
      const surroundingsViewport = { width: bounds.width, height: bounds.height };
      setSurroundingsSetup({
        viewBoxWidth: bounds.width,
        viewBoxHeight: bounds.height,
        placement: resolveSurroundingsPlacement(surroundingsBounds, surroundingsViewport),
        nodes: AVATAR_MOTION_V2.surroundingsV2
          ? createSurroundingsNodes(
            motionSeedKey,
            surroundingsBounds,
            surroundingsViewport,
            isCompactViewport
          )
          : [],
        twinkles: AVATAR_MOTION_V2.twinklesV1
          ? createTwinkleNodes(
            motionSeedKey,
            buildTwinkleAnchors(
              renderedVisage,
              surroundingsBounds,
              surroundingsViewport,
              haloBounds,
              pageLinesBounds,
              pageStackBounds,
              pagesBounds
            ),
            isCompactViewport
          )
          : []
      });
    } else {
      setSurroundingsSetup(undefined);
    }
  }, [inlineSvgMarkup, isCompactViewport, motionSeedKey, profile, renderedVisage, setAvatarRenderState, tintPalette]);

  React.useEffect(() => {
    surroundingsNodeRefs.current = surroundingsNodeRefs.current.slice(0, surroundingsSetup?.nodes.length ?? 0);
    surroundingsTwinkleRefs.current = surroundingsTwinkleRefs.current.slice(0, surroundingsSetup?.twinkles.length ?? 0);
  }, [surroundingsSetup?.nodes.length, surroundingsSetup?.twinkles.length]);

  React.useEffect(() => {
    const speechMouthAnalyzer = speechMouthRef.current;
    const shouldUseRealtimeSpeechMouth = !!remoteStream;
    if (shouldUseRealtimeSpeechMouth) {
      speechMouthAnalyzer.connect(remoteStream);
    } else {
      speechMouthAnalyzer.disconnect();
    }
    return () => speechMouthAnalyzer.disconnect();
  }, [remoteStream]);

  React.useEffect(() => {
    latestActionCueRef.current = actionCue;
  }, [actionCue]);

  React.useEffect(() => {
    motionProfileRef.current = new AvatarMotionProfile(motionSeedKey);
  }, [motionSeedKey]);

  React.useEffect(() => {
    if (isDocumentHidden) {
      frameRef.current = 0;
      return undefined;
    }

    let phase = 0;
    let lastRenderedAtMs = 0;
    const IDLE_FRAME_INTERVAL_MS = 1000 / 30;
    const animate = (): void => {
      const nowMs = Date.now();
      const bindings = bindingsRef.current;
      if (isActive) {
        const motionProfile = motionProfileRef.current;
        const storeCue = useGrimoireStore.getState().avatarActionCue;
        const incomingCue = storeCue || latestActionCueRef.current;
        if (incomingCue && incomingCue.id !== activeActionCueIdRef.current) {
          activeActionCueRef.current = incomingCue;
          activeActionCueIdRef.current = incomingCue.id;
          missingBindingsCueIdRef.current = undefined;
          const hasBindings = !!bindingsRef.current;
          logService.info(
            'system',
            `Avatar cue received: ${incomingCue.type} (#${incomingCue.id})`,
            `visage=${renderedVisage}, hasBindings=${hasBindings}, svgReady=${inlineSvgMarkup.length > 0}`
          );
        }

        const speechMouth = remoteStream ? speechMouthRef.current.sample() : undefined;
        const realSpeechActive = (speechMouth?.openness ?? 0) >= 0.05;

        let cueEyeOffsetXAdd = 0;
        let cueRootScaleX = 1;
        let cueRootScaleY = 1;
        let cueRootYOffsetAdd = 0;
        let cueMouthOverride: IAvatarCueMouthParams | undefined;

        const activeCue = activeActionCueRef.current;
        if (activeCue) {
          const cueFrame = evaluateAvatarActionCue(activeCue.type, nowMs - activeCue.at, realSpeechActive);
          cueEyeOffsetXAdd = cueFrame.eyeOffsetXAdd;
          cueRootScaleX = cueFrame.rootScaleX;
          cueRootScaleY = cueFrame.rootScaleY;
          cueRootYOffsetAdd = cueFrame.rootYOffsetAdd;
          cueMouthOverride = cueFrame.mouthOverride;
          if (cueFrame.finished) {
            logService.debug('system', `Avatar cue finished: ${activeCue.type} (#${activeCue.id})`);
            recentCueTrailRef.current = {
              type: activeCue.type,
              endedAtMs: nowMs
            };
            activeActionCueRef.current = undefined;
          }
        }

        let sampled: IMouthParams = speechMouth || { openness: 0, width: 0.5, round: 0 };
        if (cueMouthOverride) {
          sampled = blendMouthParams(sampled, cueMouthOverride, 0.82);
        }

        const shouldThrottleIdleFrame = (
          !realSpeechActive
          && !activeCue
          && effectiveExpression === 'idle'
          && assistantPlaybackState !== 'playing'
        );
        if (shouldThrottleIdleFrame && lastRenderedAtMs > 0 && (nowMs - lastRenderedAtMs) < IDLE_FRAME_INTERVAL_MS) {
          frameRef.current = requestAnimationFrame(animate);
          return;
        }
        lastRenderedAtMs = nowMs;

        if (bindings) {
          if (!firstAnimatedFrameReadyRef.current) {
            firstAnimatedFrameReadyRef.current = true;
            firstAnimatedMetricPendingRef.current = false;
            setAvatarRenderState('animated-ready');
            completeStartupMetric('first-animated-frame', `visage=${renderedVisage}`);
          }
          const recentCueTrail = recentCueTrailRef.current;
          const cueState = {
            activeType: activeCue?.type,
            activeElapsedMs: activeCue ? nowMs - activeCue.at : undefined,
            recentType: recentCueTrail?.type,
            recentElapsedMs: recentCueTrail ? nowMs - recentCueTrail.endedAtMs : undefined
          };
          const surroundingsIntensity = (AVATAR_MOTION_V2.ambientAuraV1 || AVATAR_MOTION_V2.surroundingsV2)
            ? resolveSurroundingsIntensity(effectiveExpression, cueState, realSpeechActive)
            : undefined;
          const twinkleIntensity = AVATAR_MOTION_V2.twinklesV1
            ? resolveTwinkleIntensity(effectiveExpression, cueState, realSpeechActive)
            : undefined;
          const expressionModel = resolveExpression(effectiveExpression);
          const factor = bindings.boundsWidth / 512;
          const mouthTransform = computeMouthTransform({
            openness: sampled.openness,
            width: sampled.width,
            round: sampled.round,
            mouthLift: expressionModel.mouthLift,
            mouthWidthBoost: expressionModel.mouthWidthBoost,
            mouthOpenBoost: expressionModel.mouthOpenBoost,
            factor
          }, AVATAR_MOTION_V2.speechMouthV2);
          const refinedMouthTransform = refineMouthTransformForVisage(renderedVisage, mouthTransform);
          const mouthOpacity = resolveMouthOpacityForVisage(renderedVisage, refinedMouthTransform);

          const browOffset = expressionModel.browOffset * factor;
          const gazeAnchor = gazeTarget === 'action-panel' ? 8.2 : 0;
          const eyeJitterAmplitude = motionProfile?.getEyeJitterAmplitude() ?? 1;
          const s = Math.sin(phase * 0.7);
          const c = Math.cos(phase * 0.61);
          const eyeOffsetX = (gazeAnchor + (s * 1.75 * eyeJitterAmplitude) + cueEyeOffsetXAdd) * factor;
          const eyeOffsetY = (c * 1.18) * factor;

          const blinkAmount = AVATAR_MOTION_V2.blinkV2 ? (motionProfile?.sampleBlink(nowMs) ?? 0) : 0;
          const eyeScale = AVATAR_MOTION_V2.blinkV2
            ? composeEyeScale(expressionModel.eyeScale, blinkAmount, realSpeechActive)
            : expressionModel.eyeScale;
          const refinedEyeMotion = resolveEyeMotionForVisage(renderedVisage, eyeOffsetX, eyeOffsetY, eyeScale, factor, phase);
          const refinedBrowMotion = refineBrowMotionForVisage(renderedVisage, browOffset, expressionModel.browRotation);

          const ambientSample = AVATAR_MOTION_V2.ambientV2 ? motionProfile?.sampleAmbient(nowMs) : undefined;
          const rootPulseAmp = (motionProfile?.getRootPulseAmplitude() ?? 0.006) * 1.9;
          const rootFloatAmp = (motionProfile?.getRootFloatAmplitude() ?? 1.8) * 1.7;
          const rootAmbientScale = 1 + (((ambientSample?.scale ?? 1) - 1) * 1.8);
          const rootPulse = 1 + (Math.sin(phase * 0.31) * rootPulseAmp);
          const rootScaleX = rootPulse * rootAmbientScale * cueRootScaleX;
          const rootScaleY = rootPulse * rootAmbientScale * cueRootScaleY;
          const rootFloatX = ((ambientSample?.x ?? 0) * 1.55) * factor;
          const rootFloatY = (
            Math.sin(phase * 0.45) * (rootFloatAmp * factor)
            + (((ambientSample?.y ?? 0) * 1.2) * factor)
            + refinedMouthTransform.jawOffsetY
            + (cueRootYOffsetAdd * factor)
          );

          applyTransform(bindings.mouth, 0, refinedMouthTransform.liftY, refinedMouthTransform.scaleX, refinedMouthTransform.scaleY, 0);
          applyOpacity(bindings.mouth, mouthOpacity);
          applyTransform(bindings.brows, 0, refinedBrowMotion.browOffset, 1, 1, refinedBrowMotion.browRotation);

          if (bindings.leftEye || bindings.rightEye) {
            applyTransform(bindings.leftEye, refinedEyeMotion.leftEyeOffsetX, refinedEyeMotion.leftEyeOffsetY, 1, refinedEyeMotion.eyeScale, 0);
            applyTransform(bindings.rightEye, refinedEyeMotion.rightEyeOffsetX, refinedEyeMotion.rightEyeOffsetY, 1, refinedEyeMotion.eyeScale, 0);
          } else {
            const combinedEyeOffsetX = (refinedEyeMotion.leftEyeOffsetX + refinedEyeMotion.rightEyeOffsetX) / 2;
            const combinedEyeOffsetY = (refinedEyeMotion.leftEyeOffsetY + refinedEyeMotion.rightEyeOffsetY) / 2;
            applyTransform(bindings.eyes, combinedEyeOffsetX, combinedEyeOffsetY, 1, refinedEyeMotion.eyeScale, 0);
          }

          applyTransform(bindings.faceRoot, rootFloatX, rootFloatY, rootScaleX, rootScaleY, 0);

          applyTransform(bindings.accentParticles, 0, 0, 1, 1, 0);
          applyOpacity(bindings.accentParticles, 0);
          resetClassicParallax(bindings);

          if ((AVATAR_MOTION_V2.ambientAuraV1 || AVATAR_MOTION_V2.surroundingsV2) && surroundingsSetup) {
            const intensity = surroundingsIntensity ?? resolveSurroundingsIntensity(effectiveExpression, cueState, realSpeechActive);
            const surroundingsViewport = {
              width: surroundingsSetup.viewBoxWidth,
              height: surroundingsSetup.viewBoxHeight
            };
            const auraElement = surroundingsAuraRef.current;
            if (auraElement) {
              const auraScale = 1 + (intensity.total * 0.08) + (Math.sin(phase * 0.24) * 0.02);
              const auraOpacity = 0.015 + (intensity.total * 0.07);
              const { centerX, centerY } = surroundingsSetup.placement;
              auraElement.setAttribute(
                'transform',
                `translate(${centerX.toFixed(3)} ${centerY.toFixed(3)}) scale(${auraScale.toFixed(3)}) translate(${(-centerX).toFixed(3)} ${(-centerY).toFixed(3)})`
              );
              auraElement.style.opacity = `${clamp01(auraOpacity)}`;
            }

            if (AVATAR_MOTION_V2.surroundingsV2 && surroundingsSetup.nodes.length > 0) {
              for (let index = 0; index < surroundingsSetup.nodes.length; index++) {
                const node = surroundingsSetup.nodes[index];
                const element = surroundingsNodeRefs.current[index];
                if (!element) continue;

                const frame = sampleSurroundingsNodeFrame(node, nowMs, intensity.total, surroundingsViewport);
                element.setAttribute(
                  'transform',
                  `translate(${frame.x.toFixed(3)} ${frame.y.toFixed(3)}) scale(${frame.scale.toFixed(3)})`
                );
                element.style.opacity = `${frame.opacity}`;
              }
            } else {
              applySvgElementsOpacity(surroundingsNodeRefs.current, 0);
            }
          } else {
            applySvgElementOpacity(surroundingsAuraRef.current, 0);
            applySvgElementsOpacity(surroundingsNodeRefs.current, 0);
          }

          if (AVATAR_MOTION_V2.twinklesV1 && surroundingsSetup && surroundingsSetup.twinkles.length > 0) {
            const intensity = twinkleIntensity ?? resolveTwinkleIntensity(effectiveExpression, cueState, realSpeechActive);
            for (let index = 0; index < surroundingsSetup.twinkles.length; index++) {
              const twinkle = surroundingsSetup.twinkles[index];
              const element = surroundingsTwinkleRefs.current[index];
              if (!element) continue;

              const frame = sampleTwinkleFrame(twinkle, nowMs, intensity.total);
              element.setAttribute(
                'transform',
                `translate(${frame.x.toFixed(3)} ${frame.y.toFixed(3)}) scale(${frame.scale.toFixed(3)})`
              );
              element.style.opacity = `${frame.opacity}`;
            }
          } else {
            applySvgElementsOpacity(surroundingsTwinkleRefs.current, 0);
          }

        } else {
          // No bindings yet (SVG still loading). Skip this frame.
          applySvgElementOpacity(surroundingsAuraRef.current, 0);
          applySvgElementsOpacity(surroundingsNodeRefs.current, 0);
          applySvgElementsOpacity(surroundingsTwinkleRefs.current, 0);
          if (activeCue && missingBindingsCueIdRef.current !== activeCue.id) {
            missingBindingsCueIdRef.current = activeCue.id;
            logService.warning(
              'system',
              `Avatar cue active without bindings: ${activeCue.type} (#${activeCue.id})`,
              `visage=${renderedVisage}, svgReady=${inlineSvgMarkup.length > 0}`
            );
          }
        }
      } else {
        resetClassicParallax(bindings);
        applyOpacity(bindings?.mouth, usesOpenMouthOverlay(renderedVisage) ? 0 : 1);
        applySvgElementOpacity(surroundingsAuraRef.current, 0);
        applySvgElementsOpacity(surroundingsNodeRefs.current, 0);
        applySvgElementsOpacity(surroundingsTwinkleRefs.current, 0);
      }

      phase += 0.03;
      frameRef.current = requestAnimationFrame(animate);
    };

    frameRef.current = requestAnimationFrame(animate);
    return () => {
      if (frameRef.current) {
        cancelAnimationFrame(frameRef.current);
        frameRef.current = 0;
      }
    };
  }, [
    assistantPlaybackState,
    effectiveExpression,
    gazeTarget,
    inlineSvgMarkup,
    isActive,
    isDocumentHidden,
    remoteStream,
    renderedVisage,
    setAvatarRenderState,
    surroundingsSetup
  ]);

  React.useEffect(() => {
    return () => speechMouthRef.current.disconnect();
  }, []);

  const handleMouseMove = React.useCallback((): void => {
    if (onActivity) onActivity();
  }, [onActivity]);

  return (
    <div
      onMouseMove={handleMouseMove}
      style={{
        width: width ? `${width}px` : '100%',
        height: height ? `${height}px` : '100%',
        display: 'block',
        position: 'relative',
        overflow: 'hidden',
        background: 'transparent'
      }}
    >
      <div
        style={{
          width: '100%',
          height: '100%',
          display: 'block',
          pointerEvents: 'none',
          position: 'relative',
          ...presentationTransformStyle
        }}
      >
        {surroundingsSetup && (
          <svg
            aria-hidden="true"
            viewBox={`0 0 ${surroundingsSetup.viewBoxWidth} ${surroundingsSetup.viewBoxHeight}`}
            preserveAspectRatio="xMidYMid meet"
            style={{
              position: 'absolute',
              inset: 0,
              width: '100%',
              height: '100%',
              display: 'block',
              overflow: 'visible',
              pointerEvents: 'none',
              zIndex: 0
            }}
          >
            <defs>
              <radialGradient id={auraGradientId} cx="50%" cy="50%" r="50%">
                <stop offset="0%" stopColor={tintPalette.primary} stopOpacity="0.36" />
                <stop offset="55%" stopColor={tintPalette.secondary} stopOpacity="0.16" />
                <stop offset="100%" stopColor={tintPalette.feature} stopOpacity="0" />
              </radialGradient>
            </defs>
            <ellipse
              ref={(element) => { surroundingsAuraRef.current = element || undefined; }}
              cx={surroundingsSetup.placement.centerX}
              cy={surroundingsSetup.placement.centerY}
              rx={surroundingsSetup.placement.coreRadiusX * 0.88}
              ry={surroundingsSetup.placement.coreRadiusY * 0.78}
              fill={`url(#${auraGradientId})`}
              opacity={0}
            />
          </svg>
        )}

        <div
          ref={svgHostRef}
          role="img"
          aria-label={`${renderedVisage}-avatar`}
          dangerouslySetInnerHTML={{ __html: inlineSvgMarkup }}
          style={{
            width: '100%',
            height: '100%',
            display: 'block',
            pointerEvents: 'none',
            position: 'relative',
            zIndex: 1,
            ...avatarVisualStyle
          }}
        />

        {surroundingsSetup && (surroundingsSetup.nodes.length > 0 || surroundingsSetup.twinkles.length > 0) && (
          <svg
            aria-hidden="true"
            viewBox={`0 0 ${surroundingsSetup.viewBoxWidth} ${surroundingsSetup.viewBoxHeight}`}
            preserveAspectRatio="xMidYMid meet"
            style={{
              position: 'absolute',
              inset: 0,
              width: '100%',
              height: '100%',
              display: 'block',
              overflow: 'visible',
              pointerEvents: 'none',
              zIndex: 2
            }}
          >
            {surroundingsSetup.nodes.map((node, index) => (
              <circle
                key={`surroundings-${node.id}`}
                ref={(element) => { surroundingsNodeRefs.current[index] = element || undefined; }}
                cx={0}
                cy={0}
                r={node.radius}
                fill={resolvePaletteColor(node.colorRole, tintPalette)}
                opacity={0}
              />
            ))}
            {surroundingsSetup.twinkles.map((node, index) => {
              const color = resolvePaletteColor(node.colorRole, tintPalette);
              return (
                <g
                  key={`twinkle-${node.id}`}
                  ref={(element) => { surroundingsTwinkleRefs.current[index] = element || undefined; }}
                  opacity={0}
                >
                  <line x1={-5.5} y1={0} x2={5.5} y2={0} stroke={color} strokeWidth={1.7} strokeLinecap="round" />
                  <line x1={0} y1={-5.5} x2={0} y2={5.5} stroke={color} strokeWidth={1.7} strokeLinecap="round" />
                  <line x1={-3.4} y1={-3.4} x2={3.4} y2={3.4} stroke={color} strokeWidth={0.9} strokeLinecap="round" opacity={0.64} />
                  <line x1={-3.4} y1={3.4} x2={3.4} y2={-3.4} stroke={color} strokeWidth={0.9} strokeLinecap="round" opacity={0.64} />
                  <circle cx={0} cy={0} r={1.4} fill={color} opacity={0.72} />
                </g>
              );
            })}
          </svg>
        )}
      </div>
    </div>
  );
};
