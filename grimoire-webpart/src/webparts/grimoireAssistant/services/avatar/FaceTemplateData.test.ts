import { generateFaceTemplate, generateVisageTemplate, VISAGE_OPTIONS, VisageMode } from './FaceTemplateData';
import type { IFaceTemplate } from './ParticleSystem';

function countByRegion(templates: IFaceTemplate[]): Record<string, number> {
  const result: Record<string, number> = {};
  for (const tmpl of templates) {
    result[tmpl.region] = (result[tmpl.region] || 0) + tmpl.points.length;
  }
  return result;
}

function totalPoints(templates: IFaceTemplate[]): number {
  let total = 0;
  for (const tmpl of templates) total += tmpl.points.length;
  return total;
}

function regionPoints(templates: IFaceTemplate[], region: string): Array<{ x: number; y: number }> {
  const pts: Array<{ x: number; y: number }> = [];
  for (const tmpl of templates) {
    if (tmpl.region !== region) continue;
    for (const pt of tmpl.points) pts.push(pt);
  }
  return pts;
}

function centroid(points: Array<{ x: number; y: number }>): { x: number; y: number } {
  if (points.length === 0) return { x: 0, y: 0 };
  let sx = 0;
  let sy = 0;
  for (const pt of points) {
    sx += pt.x;
    sy += pt.y;
  }
  return { x: sx / points.length, y: sy / points.length };
}

function spreadX(points: Array<{ x: number; y: number }>): number {
  if (points.length === 0) return 0;
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;
  for (const pt of points) {
    if (pt.x < min) min = pt.x;
    if (pt.x > max) max = pt.x;
  }
  return max - min;
}

function visageSignature(templates: IFaceTemplate[]): number[] {
  const leftEye = centroid(regionPoints(templates, 'left_eye'));
  const rightEye = centroid(regionPoints(templates, 'right_eye'));
  const mouthUpper = centroid(regionPoints(templates, 'mouth_upper'));
  const mouthLower = centroid(regionPoints(templates, 'mouth_lower'));
  const nose = centroid(regionPoints(templates, 'nose'));
  const ambient = regionPoints(templates, 'ambient');
  const ambientC = centroid(ambient);
  const lowerAmbient = ambient.filter((pt) => pt.y > ambientC.y + 0.06);

  return [
    leftEye.x, leftEye.y,
    rightEye.x, rightEye.y,
    mouthUpper.x, mouthUpper.y,
    mouthLower.x, mouthLower.y,
    nose.x, nose.y,
    ambientC.x, ambientC.y,
    spreadX(ambient),
    spreadX(lowerAmbient),
    mouthLower.y - mouthUpper.y
  ];
}

function euclideanDistance(a: number[], b: number[]): number {
  let sum = 0;
  for (let i = 0; i < Math.min(a.length, b.length); i++) {
    const d = a[i] - b[i];
    sum += d * d;
  }
  return Math.sqrt(sum);
}

describe('FaceTemplateData visages', () => {
  const visages = Object.keys(VISAGE_OPTIONS) as VisageMode[];
  const parametricVisages: VisageMode[] = ['classic', 'anonyMousse'];

  it('keeps identical topology and counts across all visages', () => {
    const classic = generateVisageTemplate('classic');
    const classicRegionCounts = countByRegion(classic);
    const classicTotal = totalPoints(classic);

    for (const visage of visages) {
      const template = generateVisageTemplate(visage);
      const regionCounts = countByRegion(template);
      expect(totalPoints(template)).toBe(classicTotal);
      expect(regionCounts).toEqual(classicRegionCounts);
    }
  });

  it('keeps required expressive regions populated in all visages', () => {
    const requiredRegions = [
      'left_eye',
      'right_eye',
      'left_pupil',
      'right_pupil',
      'left_eyebrow',
      'right_eyebrow',
      'mouth_upper',
      'mouth_lower',
      'nose',
      'ambient'
    ];

    for (const visage of visages) {
      const counts = countByRegion(generateVisageTemplate(visage));
      for (const region of requiredRegions) {
        expect((counts[region] || 0) > 0).toBe(true);
      }
    }
  });

  it('keeps all generated coordinates within normalized bounds', () => {
    for (const visage of visages) {
      const template = generateVisageTemplate(visage);
      for (const regionTemplate of template) {
        for (const point of regionTemplate.points) {
          expect(point.x).toBeGreaterThanOrEqual(0);
          expect(point.x).toBeLessThanOrEqual(1);
          expect(point.y).toBeGreaterThanOrEqual(0);
          expect(point.y).toBeLessThanOrEqual(1);
        }
      }
    }
  });

  it('generateFaceTemplate remains equivalent to classic visage API', () => {
    const legacy = generateFaceTemplate();
    const classic = generateVisageTemplate('classic');
    expect(totalPoints(legacy)).toBe(totalPoints(classic));
    expect(countByRegion(legacy)).toEqual(countByRegion(classic));
  });

  it('keeps visages geometrically distinct (not simple scale variants)', () => {
    const signatures = new Map<VisageMode, number[]>();
    for (const visage of parametricVisages) {
      signatures.set(visage, visageSignature(generateVisageTemplate(visage)));
    }

    for (let i = 0; i < parametricVisages.length; i++) {
      for (let j = i + 1; j < parametricVisages.length; j++) {
        const a = signatures.get(parametricVisages[i]);
        const b = signatures.get(parametricVisages[j]);
        expect(a).toBeDefined();
        expect(b).toBeDefined();
        expect(euclideanDistance(a as number[], b as number[])).toBeGreaterThan(0.03);
      }
    }
  });

  it('encodes key silhouette cues for specialized visages', () => {
    const classic = generateVisageTemplate('classic');
    const guy = generateVisageTemplate('anonyMousse');

    const classicLower = centroid(regionPoints(classic, 'mouth_lower')).y;
    const guyLower = centroid(regionPoints(guy, 'mouth_lower')).y;
    expect(guyLower).toBeGreaterThan(classicLower + 0.08);
  });

  it('keeps a soft-neutral mouth baseline across visages', () => {
    for (const visage of visages) {
      const upper = regionPoints(generateVisageTemplate(visage), 'mouth_upper')
        .slice()
        .sort((a, b) => a.x - b.x);
      expect(upper.length).toBeGreaterThan(4);

      const leftCornerY = (upper[0].y + upper[1].y) * 0.5;
      const rightCornerY = (upper[upper.length - 1].y + upper[upper.length - 2].y) * 0.5;
      const centerPoints = upper
        .slice()
        .sort((a, b) => Math.abs(a.x - 0.5) - Math.abs(b.x - 0.5))
        .slice(0, 4);
      const centerY = centroid(centerPoints).y;

      expect(leftCornerY - centerY).toBeLessThan(0.11);
      expect(rightCornerY - centerY).toBeLessThan(0.11);
    }
  });
});
