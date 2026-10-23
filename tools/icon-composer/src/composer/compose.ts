/**
 * Action Icon Composer — deterministic composition of a Main Icon plus a
 * fixed-slot Activity/Action Icon with an automatic transparent knockout.
 *
 * Output contract (single SVG document):
 *   <metadata>  composer parameters + both source SVGs (re-editable)
 *   <defs>      action-knockout-<uid>  : luminance mask (white = keep)
 *   <g>         main-icon-<uid>        : main geometry, masked
 *   <g>         action-icon-<uid>      : accent geometry at the fixed slot
 *
 * The knockout follows the *visible geometry* of the action icon dilated by
 * `knockoutPadding`. Explicit data-oe-knockout="fill" also occludes the enclosed
 * area; "bounds" describes a rectangular object whose contour is partly hidden.
 * Neither mode paints a background into the visible action.
 */

import {
  collectDrawables,
  distanceToSubpaths,
  insideSubpaths,
  matrixToAttr,
  scaleOf,
  visualBounds,
} from './geometry';
import type { Drawable, Matrix, Point } from './geometry';
import { ACTION_TYPE_ROLE, resolveAccentColor, resolveBaseColor } from './tokens';
import type { AccentRole, ComposerConfig, ComposerLayer, ComposedIcon, ComposerParams } from './types';
import { ComposerError } from './types';

const TARGET_SIZE = 24;
/** Fixed slot margin from the canvas edge, in target units (matches the references). */
const SLOT_INSET = 1;

interface ParsedSource {
  root: Element;
  viewBox: { x: number; y: number; w: number; h: number };
}

let uidCounter = 0;

const nextUid = (): string => {
  uidCounter += 1;
  return `${Date.now().toString(36)}-${uidCounter}`;
};

const parseSvg = (source: string, role: string): ParsedSource => {
  if (!source || !source.trim()) {
    throw new ComposerError(`The ${role} icon is empty. Pick or import an SVG first.`);
  }
  const doc = new DOMParser().parseFromString(source, 'image/svg+xml');
  const root = doc.documentElement;
  if (root.localName.toLowerCase() !== 'svg' || doc.querySelector('parsererror')) {
    throw new ComposerError(`The ${role} icon is not valid SVG XML.`);
  }
  const viewBoxParts = (root.getAttribute('viewBox') || '')
    .trim()
    .split(/[\s,]+/)
    .map(Number);
  const viewBox =
    viewBoxParts.length === 4 && viewBoxParts.every(Number.isFinite) && viewBoxParts[2] > 0 && viewBoxParts[3] > 0
      ? { x: viewBoxParts[0], y: viewBoxParts[1], w: viewBoxParts[2], h: viewBoxParts[3] }
      : { x: 0, y: 0, w: TARGET_SIZE, h: TARGET_SIZE };
  return { root, viewBox };
};

const isConcreteColor = (value: string | null): boolean => {
  if (!value) return false;
  const v = value.trim().toLowerCase();
  if (!v) return false;
  if (v === 'none' || v === 'transparent' || v === 'currentcolor' || v === 'inherit') return false;
  if (v.startsWith('url(')) return false;
  return true;
};

/**
 * Rewrite every hard-coded ink color to currentColor so the whole layer is
 * driven by the `color` design token set on the layer group. Colors therefore
 * never live on individual paths (oPenEfficiency two-layer icon model).
 */
const normalizeInk = (root: Element): void => {
  const walk = (el: Element): void => {
    // Mask paint is luminance data, not visible theme ink. Recoloring it destroys
    // white/black cutouts in source primitives with their own nested masks.
    if (el.tagName.toLowerCase() === 'mask') return;
    const contentSample = el.closest('[data-oe-accent-role="ContentColor"]') !== null;
    const retain = (value: string | null): boolean => contentSample && /^#(?:[\da-f]{3}|[\da-f]{6})$/i.test(value?.trim() ?? '');
    for (const attr of ['fill', 'stroke']) {
      const value = el.getAttribute(attr);
      if (isConcreteColor(value) && !retain(value)) el.setAttribute(attr, 'currentColor');
    }
    const style = el.getAttribute('style');
    if (style) {
      const rewritten = style
        .split(';')
        .map((decl) => {
          const idx = decl.indexOf(':');
          if (idx <= 0) return decl;
          const name = decl.slice(0, idx).trim().toLowerCase();
          const value = decl.slice(idx + 1).trim();
          if ((name === 'fill' || name === 'stroke') && isConcreteColor(value) && !retain(value)) {
            return `${decl.slice(0, idx)}: currentColor`;
          }
          return decl;
        })
        .join(';');
      el.setAttribute('style', rewritten);
    }
    for (const child of Array.from(el.children)) walk(child);
  };
  walk(root);
};

/**
 * Make every id inside `root` unique by suffixing it, and fix internal
 * url(#id) references accordingly. Prevents SVG id collisions when several
 * composer icons share one document.
 */
const rewriteIds = (root: Element, suffix: string): void => {
  const map = new Map<string, string>();
  const collect = (el: Element): void => {
    const id = el.getAttribute('id');
    if (id) map.set(id, `${id}-${suffix}`);
    for (const child of Array.from(el.children)) collect(child);
  };
  collect(root);
  if (map.size === 0) return;
  const apply = (el: Element): void => {
    const id = el.getAttribute('id');
    if (id && map.has(id)) el.setAttribute('id', map.get(id) as string);
    for (const attr of Array.from(el.attributes)) {
      if (!attr.value.includes('url(#')) continue;
      let value = attr.value;
      for (const [oldId, newId] of map) {
        value = value.split(`url(#${oldId})`).join(`url(#${newId})`);
      }
      el.setAttribute(attr.name, value);
    }
    for (const child of Array.from(el.children)) apply(child);
  };
  apply(root);
};

/** Presentation/root attributes that must survive when children move into a wrapper group. */
const rootPresentationAttrs = (root: Element): string => {
  const skip = new Set(['xmlns', 'width', 'height', 'viewbox', 'id', 'version', 'class', 'color']);
  const parts: string[] = [];
  for (const attr of Array.from(root.attributes)) {
    const name = attr.name.toLowerCase();
    if (skip.has(name) || name.startsWith('xmlns:')) continue;
    const value = attr.value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
    parts.push(`${attr.name}="${value}"`);
  }
  return parts.join(' ');
};

const childrenXml = (root: Element): string =>
  Array.from(root.children)
    .map((child) => new XMLSerializer().serializeToString(child))
    .join('');

const cloneSilhouette = (d: Drawable, padding: number): string => {
  const occlusion = d.element.getAttribute('data-oe-knockout');
  if (occlusion === 'bounds') {
    const bounds = visualBounds([d]);
    if (bounds) return `<rect x="${bounds.x - padding}" y="${bounds.y - padding}" width="${bounds.w + 2 * padding}" height="${bounds.h + 2 * padding}" fill="#000000"/>`;
  }
  const clone = d.element.cloneNode(false) as Element;
  clone.removeAttribute('id');
  clone.removeAttribute('style');
  clone.removeAttribute('stroke-dasharray');
  clone.removeAttribute('opacity');
  clone.setAttribute('transform', matrixToAttr(d.matrix));
  const scale = scaleOf(d.matrix);
  const baseWidth = d.hasStroke ? d.strokeWidth : 0;
  clone.setAttribute('fill', d.hasFill || occlusion === 'fill' ? '#000000' : 'none');
  clone.setAttribute('stroke', '#000000');
  clone.setAttribute('stroke-width', String(baseWidth + (2 * padding) / scale));
  clone.setAttribute('stroke-linecap', 'round');
  clone.setAttribute('stroke-linejoin', 'round');
  clone.setAttribute('fill-opacity', '1');
  clone.setAttribute('stroke-opacity', '1');
  return new XMLSerializer().serializeToString(clone);
};

/**
 * Probe points inside the transparent gap around the action icon (target
 * space). Every one of these pixels must show the background — used by the
 * visual tests to prove the knockout is truly transparent.
 *
 * Points are found on a deterministic grid: each kept point sits in the
 * middle of the transparent band (away from both the painted action surface
 * and the dilated silhouette boundary), so anti-aliasing at render time
 * cannot smear ink into the probe.
 */
const knockoutProbes = (drawables: Drawable[], padding: number, target: number): Point[] => {
  if (padding <= 1.1) return [];
  const margin = Math.min(0.5, padding * 0.25);
  const surfaceDistance = (p: Point): number => {
    let best = Number.POSITIVE_INFINITY;
    for (const d of drawables) {
      const half = d.hasStroke ? (d.strokeWidth * scaleOf(d.matrix)) / 2 : 0;
      if (d.hasFill && insideSubpaths(p, d.subpaths)) return 0;
      best = Math.min(best, Math.max(0, distanceToSubpaths(p, d.subpaths) - half));
    }
    return best;
  };
  const probes: Point[] = [];
  const step = 0.5;
  for (let x = step / 2; x < target; x += step) {
    for (let y = step / 2; y < target; y += step) {
      const dist = surfaceDistance({ x, y });
      if (dist > margin && dist < padding - margin) probes.push({ x, y });
    }
  }
  return probes;
};

const viewMatrix = (viewBox: ParsedSource['viewBox'], target: number): Matrix => {
  const scale = Math.min(target / viewBox.w, target / viewBox.h);
  return [
    scale,
    0,
    0,
    scale,
    target / 2 - (viewBox.x + viewBox.w / 2) * scale,
    target / 2 - (viewBox.y + viewBox.h / 2) * scale,
  ];
};

export interface ComposeOptions {
  /** Id suffix for this composition. Auto-generated when omitted. */
  uid?: string;
  /** Canvas size of the composed icon. Default 24. */
  targetSize?: number;
  /** Embed <metadata> for later parametric re-editing. Default true. */
  includeMetadata?: boolean;
}

export const composeIcon = (config: ComposerConfig, options: ComposeOptions = {}): ComposedIcon => {
  const uid = options.uid ?? config.uid ?? nextUid();
  const target = options.targetSize ?? TARGET_SIZE;
  const params = config.params;
  const padding = Math.max(0, params.knockoutPadding);

  const main = parseSvg(config.mainSvg, 'main');
  const action = parseSvg(config.actionSvg, 'action');
  normalizeInk(main.root);
  normalizeInk(action.root);
  rewriteIds(main.root, `m${uid}`);
  rewriteIds(action.root, `a${uid}`);

  // --- main icon: proportional viewBox fit into the target canvas ---
  const mainMatrix = viewMatrix(main.viewBox, target);
  const mainDrawables = collectDrawables(main.root, mainMatrix);
  const mainBounds = visualBounds(mainDrawables);
  if (!mainBounds || mainDrawables.length === 0) {
    throw new ComposerError('The main icon has no visible geometry.');
  }

  // --- action icon: scale by *visual* size, anchor by *visual* bounds ---
  const actionDrawablesSource = collectDrawables(action.root);
  const actionBoundsSource = visualBounds(actionDrawablesSource);
  if (!actionBoundsSource || actionDrawablesSource.length === 0) {
    throw new ComposerError('The action icon has no visible geometry.');
  }

  const mainSize = Math.max(mainBounds.w, mainBounds.h);
  const actionSourceSize = Math.max(actionBoundsSource.w, actionBoundsSource.h);
  const wantedSize = params.actionSize * mainSize;
  const scale = wantedSize / actionSourceSize;

  const inset = (SLOT_INSET * target) / TARGET_SIZE;
  const [row, column] = positionGrid(params.position);
  const anchorX = column === 'left' ? inset : column === 'right' ? target - inset : target / 2;
  const anchorY = row === 'top' ? inset : row === 'bottom' ? target - inset : target / 2;
  const sourceAnchorX =
    column === 'left'
      ? actionBoundsSource.x
      : column === 'right'
        ? actionBoundsSource.x + actionBoundsSource.w
        : actionBoundsSource.x + actionBoundsSource.w / 2;
  const sourceAnchorY =
    row === 'top'
      ? actionBoundsSource.y
      : row === 'bottom'
        ? actionBoundsSource.y + actionBoundsSource.h
        : actionBoundsSource.y + actionBoundsSource.h / 2;
  const tx = anchorX - scale * sourceAnchorX;
  const ty = anchorY - scale * sourceAnchorY;
  const actionMatrix: Matrix = [scale, 0, 0, scale, tx, ty];

  const actionDrawables = collectDrawables(action.root, actionMatrix).filter(
    (d) => d.hasFill || d.hasStroke || d.element.hasAttribute('data-oe-knockout'),
  );
  const actionBounds = visualBounds(actionDrawables);
  if (!actionBounds) {
    throw new ComposerError('The action icon has no visible geometry.');
  }

  // --- semantic accent role + colors (design tokens, never per-path literals) ---
  // Action type drives the semantic role; 'custom' falls back to the manual role.
  const accentRole: AccentRole =
    params.actionType === 'custom'
      ? params.accentRole || 'ActionBlue'
      : ACTION_TYPE_ROLE[params.actionType] || params.accentRole || 'ActionBlue';
  const accentColor = params.autoColor
    ? resolveAccentColor(accentRole, params.tokenTheme)
    : params.accentColor || resolveAccentColor(accentRole, params.tokenTheme);
  const baseColor = params.baseColor || resolveBaseColor(params.tokenTheme);

  // --- knockout mask: visible action geometry dilated by padding ---
  const silhouette = actionDrawables.map((d) => cloneSilhouette(d, padding)).join('');
  const maskId = `action-knockout-${uid}`;
  // Mask region spans well beyond the canvas so overflow geometry is never clipped.
  const m = target;
  const defs =
    `<mask id="${maskId}" data-oe-layer="action-knockout" maskUnits="userSpaceOnUse" ` +
    `x="${-m}" y="${-m}" width="${3 * m}" height="${3 * m}">` +
    `<rect x="${-m}" y="${-m}" width="${3 * m}" height="${3 * m}" fill="#ffffff"/>` +
    silhouette +
    `</mask>`;

  const mainPaint = rootPresentationAttrs(main.root);
  const actionPaint = rootPresentationAttrs(action.root);
  const mainGroup =
    `<g id="main-icon-${uid}" data-oe-layer="main-icon" mask="url(#${maskId})" color="${baseColor}">` +
    `<g transform="${matrixToAttr(mainMatrix)}"${mainPaint ? ` ${mainPaint}` : ''}>${childrenXml(main.root)}</g>` +
    `</g>`;
  const actionGroup =
    `<g id="action-icon-${uid}" data-oe-layer="action-icon" data-oe-role="accent" ` +
    `data-oe-accent-role="${accentRole}" color="${accentColor}">` +
    `<g transform="${matrixToAttr(actionMatrix)}"${actionPaint ? ` ${actionPaint}` : ''}>${childrenXml(action.root)}</g>` +
    `</g>`;

  const metadata = JSON.stringify({
    version: 1,
    uid,
    main: config.mainSvg,
    action: config.actionSvg,
    params: { ...params, accentRole } as ComposerParams,
  })
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  const buildDocument = (withMetadata: boolean): string =>
    `<svg xmlns="http://www.w3.org/2000/svg" width="${target}" height="${target}" viewBox="0 0 ${target} ${target}">` +
    (withMetadata ? `<metadata data-oe-composer="${uid}">${metadata}</metadata>` : '') +
    `<defs>${defs}</defs>${mainGroup}${actionGroup}</svg>`;

  // Editor layers stay individually editable and self-contained (each layer
  // keeps the defs it references so canvas previews render correctly).
  const layers: ComposerLayer[] = [
    {
      name: 'main-icon',
      svg:
        `<svg xmlns="http://www.w3.org/2000/svg" width="${target}" height="${target}" viewBox="0 0 ${target} ${target}">` +
        `<metadata data-oe-composer="${uid}">${metadata}</metadata>` +
        `<defs>${defs}</defs>${mainGroup}</svg>`,
    },
    {
      name: 'action-icon',
      svg:
        `<svg xmlns="http://www.w3.org/2000/svg" width="${target}" height="${target}" viewBox="0 0 ${target} ${target}">` +
        actionGroup +
        `</svg>`,
    },
  ];

  return {
    svg: buildDocument(options.includeMetadata !== false),
    flattenedSvg: buildDocument(false),
    layers,
    hasMetadata: options.includeMetadata !== false,
    mainBounds,
    actionBounds,
    probes: knockoutProbes(actionDrawables, padding, target),
    warnings: [],
  };
};

const positionGrid = (position: ComposerParams['position']): ['top' | 'middle' | 'bottom', 'left' | 'center' | 'right'] => {
  switch (position) {
    case 'top-left': return ['top', 'left'];
    case 'top-center': return ['top', 'center'];
    case 'top-right': return ['top', 'right'];
    case 'middle-left': return ['middle', 'left'];
    case 'center': return ['middle', 'center'];
    case 'middle-right': return ['middle', 'right'];
    case 'bottom-left': return ['bottom', 'left'];
    case 'bottom-center': return ['bottom', 'center'];
    case 'bottom-right': return ['bottom', 'right'];
    default: throw new ComposerError(`Unknown action position: ${String(position)}`);
  }
};

/** Read composer parameters + both sources back from an exported icon. */
export const readComposerMetadata = (svg: string): ComposerConfig | null => {
  try {
    const doc = new DOMParser().parseFromString(svg, 'image/svg+xml');
    const meta = doc.documentElement.querySelector('metadata[data-oe-composer]');
    if (!meta) return null;
    const parsed = JSON.parse(meta.textContent ?? '');
    if (typeof parsed?.main !== 'string' || typeof parsed?.action !== 'string') return null;
    return {
      mainSvg: parsed.main,
      actionSvg: parsed.action,
      params: parsed.params as ComposerParams,
      uid: typeof parsed?.uid === 'string' ? parsed.uid : undefined,
    };
  } catch {
    return null;
  }
};

/**
 * Re-import: rebuild a composition from an exported icon, restoring both
 * sources, all parameters and the id namespace (so re-imported icons never
 * collide with freshly composed ones).
 */
export const recomposeFromExport = (svg: string): ComposedIcon | null => {
  const stored = readComposerMetadata(svg);
  if (!stored) return null;
  return composeIcon(
    { mainSvg: stored.mainSvg, actionSvg: stored.actionSvg, params: stored.params },
    { uid: stored.uid },
  );
};

/** Default parameters — bottom-right slot, 40% action size, 2 unit transparent gap. */
export const defaultParams = (): ComposerParams => ({
  actionType: 'add',
  position: 'bottom-right',
  actionSize: 0.4,
  knockoutPadding: 2,
  autoColor: true,
  accentRole: 'CreateGreen',
  accentColor: '',
  // Empty means theme-derived. It stays empty in Composer metadata so later
  // theme changes can resolve the new foreground without replacing manual ink.
  baseColor: '',
  tokenTheme: 'light',
});
