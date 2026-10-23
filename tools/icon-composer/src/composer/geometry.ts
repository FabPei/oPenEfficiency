/**
 * Deterministic SVG geometry utilities for the Action Icon Composer.
 *
 * Responsibilities:
 *  - matrix math + transform attribute parsing
 *  - path data sampling (all SVG path commands incl. arcs) for measurement
 *  - paint inheritance resolution
 *  - visual bounds (geometry incl. stroke expansion) — never just the viewBox
 *
 * Measurement uses sampled curves; emission always keeps the original
 * geometry (no flattening, no unintended deformation).
 */

export type Matrix = [number, number, number, number, number, number];

export interface Point {
  x: number;
  y: number;
}

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface Drawable {
  element: Element;
  /** Full transform: source root user space <- element local space. */
  matrix: Matrix;
  hasFill: boolean;
  hasStroke: boolean;
  /** Resolved stroke width in element-local units (0 when no stroke). */
  strokeWidth: number;
  /** Transformed sample points (source root user space), used for bounds. */
  points: Point[];
  /** Transformed subpaths (source root user space), used for fill tests. */
  subpaths: Point[][];
}

export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

export const multiply = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[2] * n[1],
  m[1] * n[0] + m[3] * n[1],
  m[0] * n[2] + m[2] * n[3],
  m[1] * n[2] + m[3] * n[3],
  m[0] * n[4] + m[2] * n[5] + m[4],
  m[1] * n[4] + m[3] * n[5] + m[5],
];

export const applyPoint = (m: Matrix, x: number, y: number): Point => ({
  x: m[0] * x + m[2] * y + m[4],
  y: m[1] * x + m[3] * y + m[5],
});

/** Uniform scale factor of a matrix (geometric mean of the axis scales). */
export const scaleOf = (m: Matrix): number => {
  const det = Math.abs(m[0] * m[3] - m[1] * m[2]);
  return det > 1e-12 ? Math.sqrt(det) : 1e-6;
};

const TRANSFORM_RE = /([a-zA-Z]+)\s*\(([^)]*)\)/g;

export const parseTransform = (value: string | null): Matrix => {
  if (!value) return IDENTITY;
  let result: Matrix = IDENTITY;
  TRANSFORM_RE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = TRANSFORM_RE.exec(value)) !== null) {
    const name = match[1].toLowerCase();
    const args = match[2]
      .split(/[\s,]+/)
      .filter((part) => part.length > 0)
      .map(Number);
    if (args.some((n) => !Number.isFinite(n))) continue;
    let next: Matrix = IDENTITY;
    switch (name) {
      case 'matrix':
        if (args.length >= 6) next = [args[0], args[1], args[2], args[3], args[4], args[5]];
        break;
      case 'translate':
        next = [1, 0, 0, 1, args[0] ?? 0, args[1] ?? 0];
        break;
      case 'scale':
        next = [args[0] ?? 1, 0, 0, args.length > 1 ? args[1] : (args[0] ?? 1), 0, 0];
        break;
      case 'rotate': {
        const angle = ((args[0] ?? 0) * Math.PI) / 180;
        const cos = Math.cos(angle);
        const sin = Math.sin(angle);
        const rot: Matrix = [cos, sin, -sin, cos, 0, 0];
        if (args.length >= 3) {
          next = multiply(multiply([1, 0, 0, 1, args[1], args[2]], rot), [
            1, 0, 0, 1, -args[1], -args[2],
          ]);
        } else {
          next = rot;
        }
        break;
      }
      case 'skewx': {
        const t = Math.tan(((args[0] ?? 0) * Math.PI) / 180);
        next = [1, 0, t, 1, 0, 0];
        break;
      }
      case 'skewy': {
        const t = Math.tan(((args[0] ?? 0) * Math.PI) / 180);
        next = [1, t, 0, 1, 0, 0];
        break;
      }
      default:
        break;
    }
    result = multiply(result, next);
  }
  return result;
};

export const matrixToAttr = (m: Matrix): string =>
  `matrix(${m.map((n) => (Number.isInteger(n) ? String(n) : Number(n.toFixed(6)).toString())).join(' ')})`;

// ---------------------------------------------------------------------------
// Path sampling
// ---------------------------------------------------------------------------

const tokenizePath = (d: string): Array<{ type: 'cmd'; value: string } | { type: 'num'; value: number }> => {
  const tokens: Array<{ type: 'cmd'; value: string } | { type: 'num'; value: number }> = [];
  const re = /([MmLlHhVvCcSsQqTtAaZz])|(-?(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?)/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(d)) !== null) {
    if (match[1]) tokens.push({ type: 'cmd', value: match[1] });
    else tokens.push({ type: 'num', value: Number(match[2]) });
  }
  return tokens;
};

const CURVE_SAMPLES = 24;

/**
 * Flatten SVG path data into sampled polylines (local coordinates).
 * Supports every path command; sampling density is fixed so bounds are
 * deterministic across runs.
 */
export const samplePathData = (d: string): Array<{ points: Point[]; closed: boolean }> => {
  const tokens = tokenizePath(d);
  const subpaths: Array<{ points: Point[]; closed: boolean }> = [];
  let current: { points: Point[]; closed: boolean } = { points: [], closed: false };
  let cursor: Point = { x: 0, y: 0 };
  let start: Point = { x: 0, y: 0 };
  let lastCmd = '';
  let lastControl: Point | null = null;
  let index = 0;

  const push = (p: Point) => {
    current.points.push(p);
  };
  const beginSubpath = (p: Point) => {
    if (current.points.length > 0) subpaths.push(current);
    current = { points: [p], closed: false };
    start = p;
    cursor = p;
  };
  const num = (): number => {
    const token = tokens[index];
    if (token && token.type === 'num') {
      index += 1;
      return token.value;
    }
    return Number.NaN;
  };

  while (index < tokens.length) {
    const token = tokens[index];
    if (token.type !== 'cmd') {
      index += 1;
      continue;
    }
    index += 1;
    const cmd = token.value;
    const rel = cmd === cmd.toLowerCase() && cmd !== 'Z' && cmd !== 'z';
    const upper = cmd.toUpperCase();
    const base = (axis: 'x' | 'y'): number => (rel ? (axis === 'x' ? cursor.x : cursor.y) : 0);

    switch (upper) {
      case 'M': {
        const x = num() + base('x');
        const y = num() + base('y');
        if (!Number.isFinite(x) || !Number.isFinite(y)) break;
        beginSubpath({ x, y });
        lastControl = null;
        // subsequent coordinate pairs are implicit lineto
        while (index < tokens.length && tokens[index].type === 'num') {
          const lx = num() + base('x');
          const ly = num() + base('y');
          if (!Number.isFinite(lx) || !Number.isFinite(ly)) break;
          push({ x: lx, y: ly });
          cursor = { x: lx, y: ly };
        }
        break;
      }
      case 'L': {
        while (index < tokens.length && tokens[index].type === 'num') {
          const x = num() + base('x');
          const y = num() + base('y');
          if (!Number.isFinite(x) || !Number.isFinite(y)) break;
          push({ x, y });
          cursor = { x, y };
        }
        lastControl = null;
        break;
      }
      case 'H': {
        while (index < tokens.length && tokens[index].type === 'num') {
          const x = num() + base('x');
          if (!Number.isFinite(x)) break;
          push({ x, y: cursor.y });
          cursor = { x, y: cursor.y };
        }
        lastControl = null;
        break;
      }
      case 'V': {
        while (index < tokens.length && tokens[index].type === 'num') {
          const y = num() + base('y');
          if (!Number.isFinite(y)) break;
          push({ x: cursor.x, y });
          cursor = { x: cursor.x, y };
        }
        lastControl = null;
        break;
      }
      case 'C': {
        while (index < tokens.length && tokens[index].type === 'num') {
          const x1 = num() + base('x');
          const y1 = num() + base('y');
          const x2 = num() + base('x');
          const y2 = num() + base('y');
          const x = num() + base('x');
          const y = num() + base('y');
          if ([x1, y1, x2, y2, x, y].some((n) => !Number.isFinite(n))) break;
          sampleCubic(cursor, { x: x1, y: y1 }, { x: x2, y: y2 }, { x, y }, push);
          lastControl = { x: x2, y: y2 };
          cursor = { x, y };
        }
        break;
      }
      case 'S': {
        while (index < tokens.length && tokens[index].type === 'num') {
          const x2 = num() + base('x');
          const y2 = num() + base('y');
          const x = num() + base('x');
          const y = num() + base('y');
          if ([x2, y2, x, y].some((n) => !Number.isFinite(n))) break;
          const reflected =
            lastCmd === 'C' || lastCmd === 'S'
              ? { x: 2 * cursor.x - (lastControl?.x ?? cursor.x), y: 2 * cursor.y - (lastControl?.y ?? cursor.y) }
              : cursor;
          sampleCubic(cursor, reflected, { x: x2, y: y2 }, { x, y }, push);
          lastControl = { x: x2, y: y2 };
          cursor = { x, y };
          lastCmd = 'S';
        }
        break;
      }
      case 'Q': {
        while (index < tokens.length && tokens[index].type === 'num') {
          const x1 = num() + base('x');
          const y1 = num() + base('y');
          const x = num() + base('x');
          const y = num() + base('y');
          if ([x1, y1, x, y].some((n) => !Number.isFinite(n))) break;
          sampleQuadratic(cursor, { x: x1, y: y1 }, { x, y }, push);
          lastControl = { x: x1, y: y1 };
          cursor = { x, y };
        }
        break;
      }
      case 'T': {
        while (index < tokens.length && tokens[index].type === 'num') {
          const x = num() + base('x');
          const y = num() + base('y');
          if (![x, y].every((n) => Number.isFinite(n))) break;
          const reflected: Point =
            lastCmd === 'Q' || lastCmd === 'T'
              ? { x: 2 * cursor.x - (lastControl?.x ?? cursor.x), y: 2 * cursor.y - (lastControl?.y ?? cursor.y) }
              : cursor;
          sampleQuadratic(cursor, reflected, { x, y }, push);
          lastControl = reflected;
          cursor = { x, y };
          lastCmd = 'T';
        }
        break;
      }
      case 'A': {
        while (index < tokens.length && tokens[index].type === 'num') {
          const rx = Math.abs(num());
          const ry = Math.abs(num());
          const rotation = num();
          // arc flags are single digits and may be written without separators
          const rawFlag = tokens[index];
          let largeArc = 0;
          let sweep = 0;
          if (rawFlag && rawFlag.type === 'num') {
            index += 1;
            largeArc = rawFlag.value ? 1 : 0;
          }
          const rawSweep = tokens[index];
          if (rawSweep && rawSweep.type === 'num') {
            index += 1;
            sweep = rawSweep.value ? 1 : 0;
          }
          const x = num() + base('x');
          const y = num() + base('y');
          if (![rx, ry, rotation, x, y].every((n) => Number.isFinite(n))) break;
          sampleArc(cursor, rx, ry, rotation, largeArc, sweep, { x, y }, push);
          cursor = { x, y };
          lastControl = null;
        }
        break;
      }
      case 'Z': {
        current.closed = true;
        cursor = start;
        lastControl = null;
        break;
      }
      default:
        break;
    }
    lastCmd = upper;
  }
  if (current.points.length > 0) subpaths.push(current);
  return subpaths;
};

const sampleCubic = (
  p0: Point,
  p1: Point,
  p2: Point,
  p3: Point,
  push: (p: Point) => void,
): void => {
  for (let i = 1; i <= CURVE_SAMPLES; i += 1) {
    const t = i / CURVE_SAMPLES;
    const mt = 1 - t;
    push({
      x: mt * mt * mt * p0.x + 3 * mt * mt * t * p1.x + 3 * mt * t * t * p2.x + t * t * t * p3.x,
      y: mt * mt * mt * p0.y + 3 * mt * mt * t * p1.y + 3 * mt * t * t * p2.y + t * t * t * p3.y,
    });
  }
};

const sampleQuadratic = (p0: Point, p1: Point, p2: Point, push: (p: Point) => void): void => {
  for (let i = 1; i <= CURVE_SAMPLES; i += 1) {
    const t = i / CURVE_SAMPLES;
    const mt = 1 - t;
    push({
      x: mt * mt * p0.x + 2 * mt * t * p1.x + t * t * p2.x,
      y: mt * mt * p0.y + 2 * mt * t * p1.y + t * t * p2.y,
    });
  }
};

const sampleArc = (
  p0: Point,
  rx: number,
  ry: number,
  rotationDeg: number,
  largeArc: number,
  sweep: number,
  p1: Point,
  push: (p: Point) => void,
): void => {
  if (rx === 0 || ry === 0) {
    push(p1);
    return;
  }
  const phi = (rotationDeg * Math.PI) / 180;
  const cosPhi = Math.cos(phi);
  const sinPhi = Math.sin(phi);
  const dx2 = (p0.x - p1.x) / 2;
  const dy2 = (p0.y - p1.y) / 2;
  const x1p = cosPhi * dx2 + sinPhi * dy2;
  const y1p = -sinPhi * dx2 + cosPhi * dy2;
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  const scale = lambda > 1 ? Math.sqrt(lambda) : 1;
  const rxS = rx * scale;
  const ryS = ry * scale;
  const sign = largeArc !== sweep ? 1 : -1;
  const denom = rxS * rxS * y1p * y1p + ryS * ryS * x1p * x1p;
  const numer = Math.max(0, rxS * rxS * ryS * ryS - denom);
  // c' = ± (rx·y'/ry, -ry·x'/rx) · sqrt(numer/denom)   (SVG endpoint→center spec)
  const coef = denom === 0 ? 0 : sign * Math.sqrt(numer / denom);
  const cxp = (coef * rxS * y1p) / ryS;
  const cyp = (-coef * ryS * x1p) / rxS;
  const cx = cosPhi * cxp - sinPhi * cyp + (p0.x + p1.x) / 2;
  const cy = sinPhi * cxp + cosPhi * cyp + (p0.y + p1.y) / 2;
  const angle = (ux: number, uy: number, vx: number, vy: number): number => {
    const dot = ux * vx + uy * vy;
    const len = Math.sqrt((ux * ux + uy * uy) * (vx * vx + vy * vy));
    const raw = len === 0 ? 0 : Math.acos(Math.min(1, Math.max(-1, dot / len)));
    return ux * vy - uy * vx < 0 ? -raw : raw;
  };
  const theta1 = angle(1, 0, (x1p - cxp) / rxS, (y1p - cyp) / ryS);
  let delta = angle((x1p - cxp) / rxS, (y1p - cyp) / ryS, (-x1p - cxp) / rxS, (-y1p - cyp) / ryS);
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  if (sweep && delta < 0) delta += 2 * Math.PI;
  const steps = Math.max(2, Math.ceil((Math.abs(delta) / Math.PI) * CURVE_SAMPLES));
  for (let i = 1; i <= steps; i += 1) {
    const theta = theta1 + (delta * i) / steps;
    push({
      x: cx + rxS * Math.cos(theta) * cosPhi - ryS * Math.sin(theta) * sinPhi,
      y: cy + rxS * Math.cos(theta) * sinPhi + ryS * Math.sin(theta) * cosPhi,
    });
  }
};

// ---------------------------------------------------------------------------
// Paint resolution + drawable collection
// ---------------------------------------------------------------------------

interface PaintState {
  fill: string;
  stroke: string;
  strokeWidth: number;
}

const DEFAULT_PAINT: PaintState = { fill: '#000000', stroke: 'none', strokeWidth: 1 };

const parseStyle = (el: Element): Record<string, string> => {
  const out: Record<string, string> = {};
  const style = el.getAttribute('style');
  if (!style) return out;
  for (const decl of style.split(';')) {
    const idx = decl.indexOf(':');
    if (idx <= 0) continue;
    out[decl.slice(0, idx).trim().toLowerCase()] = decl.slice(idx + 1).trim();
  }
  return out;
};

const paintOf = (el: Element, style: Record<string, string>, name: string): string | null =>
  style[name] ?? el.getAttribute(name);

export const resolvePaint = (el: Element, inherited: PaintState): PaintState => {
  const style = parseStyle(el);
  const fill = paintOf(el, style, 'fill');
  const stroke = paintOf(el, style, 'stroke');
  const strokeWidthRaw = paintOf(el, style, 'stroke-width');
  const strokeWidth = strokeWidthRaw ? Number.parseFloat(strokeWidthRaw) : Number.NaN;
  return {
    fill: fill ?? inherited.fill,
    stroke: stroke ?? inherited.stroke,
    strokeWidth: Number.isFinite(strokeWidth) && strokeWidth >= 0 ? strokeWidth : inherited.strokeWidth,
  };
};

const shapeSubpaths = (el: Element): Point[][] => {
  const tag = el.localName.toLowerCase();
  const n = (name: string): number => Number.parseFloat(el.getAttribute(name) ?? '0');
  switch (tag) {
    case 'path':
      return samplePathData(el.getAttribute('d') ?? '').map((sp) => sp.points);
    case 'rect': {
      const x = n('x');
      const y = n('y');
      const w = n('width');
      const h = n('height');
      return [
        [
          { x, y },
          { x: x + w, y },
          { x: x + w, y: y + h },
          { x, y: y + h },
        ],
      ];
    }
    case 'circle': {
      const cx = n('cx');
      const cy = n('cy');
      const r = n('r');
      const pts: Point[] = [];
      for (let i = 0; i < 32; i += 1) {
        const a = (i / 32) * Math.PI * 2;
        pts.push({ x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) });
      }
      return [pts];
    }
    case 'ellipse': {
      const cx = n('cx');
      const cy = n('cy');
      const rx = n('rx');
      const ry = n('ry');
      const pts: Point[] = [];
      for (let i = 0; i < 32; i += 1) {
        const a = (i / 32) * Math.PI * 2;
        pts.push({ x: cx + rx * Math.cos(a), y: cy + ry * Math.sin(a) });
      }
      return [pts];
    }
    case 'line':
      return [
        [
          { x: n('x1'), y: n('y1') },
          { x: n('x2'), y: n('y2') },
        ],
      ];
    case 'polyline':
    case 'polygon': {
      const raw = (el.getAttribute('points') ?? '').trim();
      const nums = raw
        .split(/[\s,]+/)
        .filter((p) => p.length > 0)
        .map(Number)
        .filter((v) => Number.isFinite(v));
      const pts: Point[] = [];
      for (let i = 0; i + 1 < nums.length; i += 2) pts.push({ x: nums[i], y: nums[i + 1] });
      return pts.length > 0 ? [pts] : [];
    }
    default:
      return [];
  }
};

const DRAWABLE_TAGS = new Set(['path', 'rect', 'circle', 'ellipse', 'line', 'polyline', 'polygon']);
const CONTAINER_TAGS = new Set(['svg', 'g', 'a']);

const isHidden = (el: Element): boolean => {
  const style = parseStyle(el);
  const display = style.display ?? el.getAttribute('display');
  const visibility = style.visibility ?? el.getAttribute('visibility');
  return display === 'none' || visibility === 'hidden' || visibility === 'collapse';
};

/**
 * Collect all drawable geometry under `root`, resolving transforms and paint
 * inheritance. `rootMatrix` maps root user space into the desired output space.
 * Paint and transform on `root` itself are honoured (typical icon sources put
 * stroke="currentColor" stroke-width="2" on the <svg> root).
 */
export const collectDrawables = (root: Element, rootMatrix: Matrix = IDENTITY): Drawable[] => {
  const out: Drawable[] = [];
  const rootPaint = resolvePaint(root, DEFAULT_PAINT);
  const rootOwn = parseTransform(root.getAttribute('transform'));

  const walk = (el: Element, parentMatrix: Matrix, parentPaint: PaintState): void => {
    if (isHidden(el)) return;
    const tag = el.localName.toLowerCase();
    const own = parseTransform(el.getAttribute('transform'));
    const matrix = multiply(parentMatrix, own);
    const paint = resolvePaint(el, parentPaint);
    if (DRAWABLE_TAGS.has(tag)) {
      const rawSubpaths = shapeSubpaths(el);
      if (rawSubpaths.length > 0) {
        const full = multiply(rootMatrix, matrix);
        out.push({
          element: el,
          matrix: full,
          hasFill: paint.fill !== 'none' && paint.fill !== 'transparent',
          hasStroke: paint.stroke !== 'none' && paint.stroke !== 'transparent' && paint.strokeWidth > 0,
          strokeWidth: paint.strokeWidth,
          points: rawSubpaths.flat().map((p) => applyPoint(full, p.x, p.y)),
          subpaths: rawSubpaths.map((sp) => sp.map((p) => applyPoint(full, p.x, p.y))),
        });
      }
      return;
    }
    if (CONTAINER_TAGS.has(tag)) {
      for (const child of Array.from(el.children)) walk(child, matrix, paint);
    }
  };

  for (const child of Array.from(root.children)) walk(child, rootOwn, rootPaint);
  return out;
};

/** Visual bounds: geometry bounds expanded by half the effective stroke width. */
export const visualBounds = (drawables: Drawable[]): Rect | null => {
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  let seen = false;
  for (const d of drawables) {
    const scale = scaleOf(d.matrix);
    const expand = d.hasStroke ? (d.strokeWidth * scale) / 2 : 0;
    for (const p of d.points) {
      minX = Math.min(minX, p.x - expand);
      minY = Math.min(minY, p.y - expand);
      maxX = Math.max(maxX, p.x + expand);
      maxY = Math.max(maxY, p.y + expand);
      seen = true;
    }
  }
  if (!seen) return null;
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
};

/** Distance from point p to segment ab. */
export const pointSegmentDistance = (p: Point, a: Point, b: Point): number => {
  const abx = b.x - a.x;
  const aby = b.y - a.y;
  const len2 = abx * abx + aby * aby;
  if (len2 === 0) return Math.hypot(p.x - a.x, p.y - a.y);
  let t = ((p.x - a.x) * abx + (p.y - a.y) * aby) / len2;
  t = Math.max(0, Math.min(1, t));
  return Math.hypot(p.x - (a.x + t * abx), p.y - (a.y + t * aby));
};

/** Minimum distance from p to any segment of any subpath. */
export const distanceToSubpaths = (p: Point, subpaths: Point[][]): number => {
  let best = Number.POSITIVE_INFINITY;
  for (const sp of subpaths) {
    for (let i = 0; i + 1 < sp.length; i += 1) {
      best = Math.min(best, pointSegmentDistance(p, sp[i], sp[i + 1]));
    }
    if (sp.length > 2) best = Math.min(best, pointSegmentDistance(p, sp[sp.length - 1], sp[0]));
  }
  return best;
};

/** Even-odd point-in-polygon test over all subpaths (open subpaths get closed). */
export const insideSubpaths = (p: Point, subpaths: Point[][]): boolean => {
  let inside = false;
  for (const sp of subpaths) {
    for (let i = 0, j = sp.length - 1; i < sp.length; j = i, i += 1) {
      const a = sp[i];
      const b = sp[j];
      if (a.y > p.y !== b.y > p.y && p.x < ((b.x - a.x) * (p.y - a.y)) / (b.y - a.y) + a.x) {
        inside = !inside;
      }
    }
  }
  return inside;
};
