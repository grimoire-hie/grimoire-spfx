import {
  buildFallbackTwinkleAnchors,
  createTwinkleNodes,
  DESKTOP_TWINKLE_NODE_COUNT,
  MOBILE_TWINKLE_NODE_COUNT,
  resolveTwinkleIntensity,
  resolveTwinkleNodeCount,
  sampleClassicParallaxFrame,
  sampleTwinkleFrame
} from './AvatarFlourishMotion';

const viewport = { width: 560, height: 462 };
const bounds = { x: 88, y: 96, width: 384, height: 252 };

describe('AvatarFlourishMotion', () => {
  it('builds fallback twinkle anchors inside the viewport', () => {
    const anchors = buildFallbackTwinkleAnchors(bounds, viewport);

    expect(anchors.length).toBeGreaterThanOrEqual(DESKTOP_TWINKLE_NODE_COUNT);

    for (const anchor of anchors) {
      expect(anchor.x).toBeGreaterThanOrEqual(12);
      expect(anchor.x).toBeLessThanOrEqual(viewport.width - 12);
      expect(anchor.y).toBeGreaterThanOrEqual(12);
      expect(anchor.y).toBeLessThanOrEqual(viewport.height - 12);
    }
  });

  it('creates deterministic twinkle nodes and sampled frames for the same seed', () => {
    const anchors = buildFallbackTwinkleAnchors(bounds, viewport);
    const a = createTwinkleNodes('seed:user@example.com|classic', anchors, false);
    const b = createTwinkleNodes('seed:user@example.com|classic', anchors, false);

    expect(a).toEqual(b);

    const sampleA = a.slice(0, 3).map((node) => sampleTwinkleFrame(node, 3200, 0.28));
    const sampleB = b.slice(0, 3).map((node) => sampleTwinkleFrame(node, 3200, 0.28));
    expect(sampleA).toEqual(sampleB);
  });

  it('respects desktop and mobile twinkle caps', () => {
    const anchors = buildFallbackTwinkleAnchors(bounds, viewport);

    expect(resolveTwinkleNodeCount(false)).toBe(DESKTOP_TWINKLE_NODE_COUNT);
    expect(resolveTwinkleNodeCount(true)).toBe(MOBILE_TWINKLE_NODE_COUNT);
    expect(createTwinkleNodes('desktop', anchors, false)).toHaveLength(DESKTOP_TWINKLE_NODE_COUNT);
    expect(createTwinkleNodes('mobile', anchors, true)).toHaveLength(MOBILE_TWINKLE_NODE_COUNT);
  });

  it('keeps twinkle intensity clamped and cue-aware', () => {
    const idle = resolveTwinkleIntensity('idle', {}, false);
    const thinking = resolveTwinkleIntensity('thinking', {}, false);
    const speaking = resolveTwinkleIntensity('speaking', {}, true);
    const activeCue = resolveTwinkleIntensity('idle', {
      activeType: 'summarize',
      activeElapsedMs: 220
    }, false);

    expect(thinking.total).toBeGreaterThan(idle.total);
    expect(speaking.total).toBeLessThan(idle.total);
    expect(activeCue.total).toBeGreaterThan(idle.total);

    for (const sample of [idle, thinking, speaking, activeCue]) {
      expect(sample.base).toBeGreaterThanOrEqual(0);
      expect(sample.cue).toBeGreaterThanOrEqual(0);
      expect(sample.total).toBeGreaterThanOrEqual(0);
      expect(sample.total).toBeLessThanOrEqual(1);
    }
  });

  it('boosts sampled twinkle opacity during cue-weighted peaks', () => {
    const anchors = buildFallbackTwinkleAnchors(bounds, viewport);
    const [node] = createTwinkleNodes('cue-boost', anchors, false);
    const idle = resolveTwinkleIntensity('idle', {}, false);
    const cue = resolveTwinkleIntensity('idle', {
      activeType: 'focus',
      activeElapsedMs: 180
    }, false);

    let idlePeak = 0;
    let cuePeak = 0;
    for (let t = 0; t <= 9600; t += 240) {
      idlePeak = Math.max(idlePeak, sampleTwinkleFrame(node, t, idle.total).opacity);
      cuePeak = Math.max(cuePeak, sampleTwinkleFrame(node, t, cue.total).opacity);
    }

    expect(cuePeak).toBeGreaterThan(idlePeak);
    expect(cuePeak).toBeGreaterThan(0.1);
  });

  it('keeps classic parallax subtle and deterministic', () => {
    const a = sampleClassicParallaxFrame(4200, 0.52);
    const b = sampleClassicParallaxFrame(4200, 0.52);

    expect(a).toEqual(b);
    expect(Math.abs(a.haloY)).toBeLessThan(2);
    expect(Math.abs(a.linesX)).toBeLessThan(2);
    expect(a.pagesScale).toBeGreaterThan(0.99);
    expect(a.pagesScale).toBeLessThan(1.01);
  });
});
