import { IFaceTemplate, ParticleSystem, PerformanceMonitor } from './ParticleSystem';

function feedFrames(monitor: PerformanceMonitor, frameTimeMs: number, count: number): void {
  for (let i = 0; i < count; i++) {
    monitor.recordFrame(frameTimeMs);
  }
}

const TEST_TEMPLATE: IFaceTemplate[] = [
  {
    region: 'head_contour',
    points: [
      { x: 0.3, y: 0.3 },
      { x: 0.5, y: 0.2 },
      { x: 0.7, y: 0.3 }
    ]
  },
  {
    region: 'mouth_upper',
    points: [
      { x: 0.45, y: 0.65 },
      { x: 0.5, y: 0.64 },
      { x: 0.55, y: 0.65 }
    ]
  }
];

describe('PerformanceMonitor', () => {
  it('degrades to tier 3 under sustained slow frame intervals', () => {
    const monitor = new PerformanceMonitor();
    feedFrames(monitor, 26, 700);
    expect(monitor.getTier()).toBe(3);
  });

  it('recovers to tier 0 under sustained healthy frame intervals', () => {
    const monitor = new PerformanceMonitor();
    feedFrames(monitor, 26, 700);
    expect(monitor.getTier()).toBe(3);

    feedFrames(monitor, 8, 900);
    expect(monitor.getTier()).toBe(0);
  });
});

describe('ParticleSystem perf controls', () => {
  it('uses point-cloud defaults with subtle mesh', () => {
    const ps = new ParticleSystem();
    const config = ps.getConfig();

    expect(config.particleShape).toBe('circle');
    expect(config.meshStyle).toBe('subtle');
    expect(config.voxelSizeMin).toBeGreaterThan(0);
    expect(config.voxelSizeMax).toBeGreaterThan(config.voxelSizeMin);
    expect(config.voxelDepthShading).toBeGreaterThanOrEqual(0);
    expect(config.voxelDepthShading).toBeLessThanOrEqual(1);
    expect(config.voxelEdgeBreakup).toBeGreaterThanOrEqual(0);
    expect(config.voxelEdgeBreakup).toBeLessThanOrEqual(1);
  });

  it('stores normalized coordinates and pseudo-depth for template particles', () => {
    const ps = new ParticleSystem();
    ps.initFromTemplate(TEST_TEMPLATE, 800, 600, false);
    const particles = ps.getParticles();

    expect(particles.length).toBeGreaterThan(0);
    for (const p of particles) {
      expect(p.nx).toBeGreaterThanOrEqual(0);
      expect(p.nx).toBeLessThanOrEqual(1);
      expect(p.ny).toBeGreaterThanOrEqual(0);
      expect(p.ny).toBeLessThanOrEqual(1);
      expect(p.nz).toBeGreaterThanOrEqual(0);
      expect(p.nz).toBeLessThanOrEqual(1);
    }
  });

  it('applies tier-3 compatible controls without destructive core particle loss', () => {
    const ps = new ParticleSystem();
    ps.initFromTemplate(TEST_TEMPLATE, 1000, 1000, false);
    const initialCount = ps.getParticleCount();

    ps.disableMesh();
    ps.setThoughtParticleLimits(80, 1);

    expect(ps.isMeshDisabled()).toBe(true);
    expect(ps.getParticleCount()).toBe(initialCount);

    ps.enableMesh();
    expect(ps.isMeshDisabled()).toBe(false);
  });
});
