import {
  createSurroundingsNodes,
  DESKTOP_SURROUNDINGS_NODE_COUNT,
  MOBILE_SURROUNDINGS_NODE_COUNT,
  resolveSurroundingsIntensity,
  resolveSurroundingsNodeCount,
  resolveSurroundingsPlacement,
  sampleSurroundingsNodeFrame
} from './SurroundingsMotion';

const viewport = { width: 560, height: 462 };
const bounds = { x: 20, y: 78, width: 520, height: 376 };

describe('SurroundingsMotion', () => {
  it('creates deterministic node layouts and sampled frames for the same seed', () => {
    const a = createSurroundingsNodes('seed:user@example.com|classic', bounds, viewport, false);
    const b = createSurroundingsNodes('seed:user@example.com|classic', bounds, viewport, false);

    expect(a).toEqual(b);

    const sampleA = a.slice(0, 4).map((node) => sampleSurroundingsNodeFrame(node, 4321, 0.34, viewport));
    const sampleB = b.slice(0, 4).map((node) => sampleSurroundingsNodeFrame(node, 4321, 0.34, viewport));
    expect(sampleA).toEqual(sampleB);
  });

  it('respects desktop and mobile node caps', () => {
    expect(resolveSurroundingsNodeCount(false)).toBe(DESKTOP_SURROUNDINGS_NODE_COUNT);
    expect(resolveSurroundingsNodeCount(true)).toBe(MOBILE_SURROUNDINGS_NODE_COUNT);
    expect(createSurroundingsNodes('desktop', bounds, viewport, false)).toHaveLength(DESKTOP_SURROUNDINGS_NODE_COUNT);
    expect(createSurroundingsNodes('mobile', bounds, viewport, true)).toHaveLength(MOBILE_SURROUNDINGS_NODE_COUNT);
  });

  it('keeps anchors outside the core bbox band and frames inside the viewport', () => {
    const placement = resolveSurroundingsPlacement(bounds, viewport);
    const nodes = createSurroundingsNodes('bounds-check', bounds, viewport, false);

    for (const node of nodes) {
      const ellipseRatio = (
        (Math.pow(node.anchorX - placement.centerX, 2) / Math.pow(placement.coreRadiusX, 2))
        + (Math.pow(node.anchorY - placement.centerY, 2) / Math.pow(placement.coreRadiusY, 2))
      );

      expect(ellipseRatio).toBeGreaterThanOrEqual(0.98);
      expect(node.anchorX).toBeGreaterThanOrEqual(placement.viewportPadding);
      expect(node.anchorX).toBeLessThanOrEqual(viewport.width - placement.viewportPadding);
      expect(node.anchorY).toBeGreaterThanOrEqual(placement.viewportPadding);
      expect(node.anchorY).toBeLessThanOrEqual(viewport.height - placement.viewportPadding);

      for (let t = 0; t <= 4800; t += 320) {
        const frame = sampleSurroundingsNodeFrame(node, t, 0.42, viewport);
        expect(frame.x).toBeGreaterThanOrEqual(placement.viewportPadding);
        expect(frame.x).toBeLessThanOrEqual(viewport.width - placement.viewportPadding);
        expect(frame.y).toBeGreaterThanOrEqual(placement.viewportPadding);
        expect(frame.y).toBeLessThanOrEqual(viewport.height - placement.viewportPadding);
      }
    }
  });

  it('keeps intensity clamped and boosts idle/thinking states more than speaking', () => {
    const idle = resolveSurroundingsIntensity('idle', {}, false);
    const thinking = resolveSurroundingsIntensity('thinking', {}, false);
    const speaking = resolveSurroundingsIntensity('speaking', {}, true);
    const activeCue = resolveSurroundingsIntensity('idle', {
      activeType: 'summarize',
      activeElapsedMs: 240
    }, false);
    const recentCue = resolveSurroundingsIntensity('idle', {
      recentType: 'focus',
      recentElapsedMs: 120
    }, false);

    expect(idle.total).toBeGreaterThan(0);
    expect(thinking.total).toBeGreaterThan(idle.total);
    expect(speaking.total).toBeLessThan(idle.total);
    expect(speaking.total).toBeLessThanOrEqual(0.3);
    expect(activeCue.total).toBeGreaterThan(idle.total);
    expect(recentCue.cue).toBeGreaterThan(0);

    for (const sample of [idle, thinking, speaking, activeCue, recentCue]) {
      expect(sample.base).toBeGreaterThanOrEqual(0);
      expect(sample.cue).toBeGreaterThanOrEqual(0);
      expect(sample.total).toBeGreaterThanOrEqual(0);
      expect(sample.total).toBeLessThanOrEqual(1);
    }
  });

  it('keeps idle sampled nodes visibly above a near-zero opacity floor', () => {
    const [node] = createSurroundingsNodes('visible-idle', bounds, viewport, false);
    const idle = resolveSurroundingsIntensity('idle', {}, false);
    const frame = sampleSurroundingsNodeFrame(node, 1800, idle.total, viewport);

    expect(frame.opacity).toBeGreaterThan(0.18);
    expect(frame.scale).toBeGreaterThan(1);
  });
});
