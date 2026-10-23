import type { CanvasConfig, SvgElement } from '../types/svg';

export interface SanitizedSvg {
  content: string;
  width: number;
  height: number;
  viewBox: { x: number; y: number; w: number; h: number };
}

export const pathTransformFor = (element: SvgElement): string => {
  const box = element.pathViewBox ?? { x: 0, y: 0, w: element.width || 1, h: element.height || 1 };
  return `translate(${element.x} ${element.y}) scale(${element.width / box.w} ${element.height / box.h}) translate(${-box.x} ${-box.y})`;
};

const blockedElements = new Set([
  'script',
  'foreignobject',
  'iframe',
  'object',
  'embed',
  'audio',
  'video',
  'canvas',
  'link',
  'meta',
  'base',
  'style',
  'animate',
  'animatemotion',
  'animatetransform',
  'set',
  'discard',
]);

const numericValue = (value: string | null): number | null => {
  if (!value) return null;
  if (!/^\s*(?:\d+\.?\d*|\.\d+)(?:px)?\s*$/i.test(value)) return null;
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
};

const safeUrlValue = (value: string, tagName: string): boolean => {
  const normalized = value.trim().toLowerCase();
  if (normalized.startsWith('#')) return true;
  if (tagName !== 'image' && tagName !== 'feimage') return false;
  return ['data:image/png', 'data:image/jpeg', 'data:image/webp', 'data:image/svg+xml'].some(
    (prefix) => normalized.startsWith(prefix),
  );
};

export const sanitizeSvg = (source: string): SanitizedSvg => {
  const parser = new DOMParser();
  const documentNode = parser.parseFromString(source, 'image/svg+xml');
  const parseError = documentNode.querySelector('parsererror');
  const root = documentNode.documentElement;

  if (parseError || root.localName.toLowerCase() !== 'svg') {
    throw new Error('The file is not valid SVG XML.');
  }

  const allElements = [root, ...Array.from(root.querySelectorAll('*'))];
  for (const element of allElements) {
    const tagName = element.localName.toLowerCase();
    if (blockedElements.has(tagName)) {
      element.remove();
      continue;
    }

    for (const attribute of Array.from(element.attributes)) {
      const name = attribute.name.toLowerCase();
      const value = attribute.value.trim();

      if (name.startsWith('on') || name === 'src' || name === 'formaction') {
        element.removeAttribute(attribute.name);
        continue;
      }

      if ((name === 'href' || name === 'xlink:href') && !safeUrlValue(value, tagName)) {
        element.removeAttribute(attribute.name);
        continue;
      }

      const referencedUrls = Array.from(value.matchAll(/url\s*\(\s*(['"]?)(.*?)\1\s*\)/gi));
      if (referencedUrls.some((match) => !match[2].trim().startsWith('#'))) {
        element.removeAttribute(attribute.name);
        continue;
      }

      if (
        name === 'style' &&
        (/javascript\s*:/i.test(value) ||
          /expression\s*\(/i.test(value) ||
          referencedUrls.some((match) => !match[2].trim().startsWith('#')))
      ) {
        element.removeAttribute(attribute.name);
      }
    }
  }

  root.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

  const viewBoxParts = (root.getAttribute('viewBox') || '')
    .trim()
    .split(/[\s,]+/)
    .map(Number);
  const validViewBox =
    viewBoxParts.length === 4 &&
    viewBoxParts.every(Number.isFinite) &&
    viewBoxParts[2] > 0 &&
    viewBoxParts[3] > 0;

  const width = numericValue(root.getAttribute('width')) ?? (validViewBox ? viewBoxParts[2] : 24);
  const height = numericValue(root.getAttribute('height')) ?? (validViewBox ? viewBoxParts[3] : 24);
  const viewBox = validViewBox
    ? { x: viewBoxParts[0], y: viewBoxParts[1], w: viewBoxParts[2], h: viewBoxParts[3] }
    : { x: 0, y: 0, w: width, h: height };

  root.setAttribute('viewBox', `${viewBox.x} ${viewBox.y} ${viewBox.w} ${viewBox.h}`);
  return {
    content: new XMLSerializer().serializeToString(root),
    width,
    height,
    viewBox,
  };
};

export const svgToDataUri = (source: string): string =>
  `data:image/svg+xml;charset=utf-8,${encodeURIComponent(source)}`;

export const safeSvgToDataUri = (source: string): string => {
  try {
    return svgToDataUri(sanitizeSvg(source).content);
  } catch {
    return '';
  }
};

export const escapeXml = (value: string | number | undefined): string =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

const transformFor = (element: SvgElement): string => {
  const cx = element.x + element.width / 2;
  const cy = element.y + element.height / 2;
  return `translate(${cx} ${cy}) rotate(${element.rotation || 0}) scale(${element.flipX ? -1 : 1} ${element.flipY ? -1 : 1}) translate(${-cx} ${-cy})`;
};

const paintAttributes = (element: SvgElement, fill: string): string => {
  const attributes = [
    `fill="${escapeXml(fill)}"`,
    `fill-opacity="${element.fillOpacity}"`,
    `stroke="${escapeXml(element.stroke)}"`,
    `stroke-width="${element.strokeWidth}"`,
    `stroke-opacity="${element.strokeOpacity}"`,
    element.strokeDasharray ? `stroke-dasharray="${escapeXml(element.strokeDasharray)}"` : '',
    element.strokeLinecap ? `stroke-linecap="${element.strokeLinecap}"` : '',
    element.strokeLinejoin ? `stroke-linejoin="${element.strokeLinejoin}"` : '',
  ];
  return attributes.filter(Boolean).join(' ');
};

const pointsForPolygon = (element: SvgElement, sides: number): string => {
  const cx = element.x + element.width / 2;
  const cy = element.y + element.height / 2;
  return Array.from({ length: sides }, (_, index) => {
    const angle = (index * 2 * Math.PI) / sides - Math.PI / 2;
    return `${cx + (element.width / 2) * Math.cos(angle)},${cy + (element.height / 2) * Math.sin(angle)}`;
  }).join(' ');
};

const pointsForStar = (element: SvgElement): string => {
  const cx = element.x + element.width / 2;
  const cy = element.y + element.height / 2;
  const outerRadius = Math.min(element.width, element.height) / 2;
  return Array.from({ length: 10 }, (_, index) => {
    const radius = index % 2 === 0 ? outerRadius : outerRadius / 2;
    const angle = (index * Math.PI) / 5 - Math.PI / 2;
    return `${cx + radius * Math.cos(angle)},${cy + radius * Math.sin(angle)}`;
  }).join(' ');
};

const serializeRawSvg = (element: SvgElement): string => {
  if (!element.rawSvgContent) return '';
  const safe = sanitizeSvg(element.rawSvgContent);
  const parsed = new DOMParser().parseFromString(safe.content, 'image/svg+xml');
  const root = parsed.documentElement;
  const inner = Array.from(root.childNodes)
    .map((node) => new XMLSerializer().serializeToString(node))
    .join('');
  return `<g transform="${transformFor(element)}"><svg x="${element.x}" y="${element.y}" width="${element.width}" height="${element.height}" viewBox="${safe.viewBox.x} ${safe.viewBox.y} ${safe.viewBox.w} ${safe.viewBox.h}" preserveAspectRatio="xMidYMid meet">${inner}</svg></g>`;
};

export const serializeSvgDocument = (
  elements: SvgElement[],
  config: CanvasConfig,
): string => {
  const visibleElements = elements.filter((element) => element.visible);
  const definitions: string[] = [];

  for (const element of visibleElements) {
    if (element.gradient?.enabled) {
      const gradient = element.gradient;
      const id = `gradient-${escapeXml(element.id)}`;
      if (gradient.type === 'radial') {
        definitions.push(
          `<radialGradient id="${id}"><stop offset="0%" stop-color="${escapeXml(gradient.color1)}" stop-opacity="${gradient.opacity1 ?? 1}"/><stop offset="100%" stop-color="${escapeXml(gradient.color2)}" stop-opacity="${gradient.opacity2 ?? 1}"/></radialGradient>`,
        );
      } else {
        const radians = (gradient.angle * Math.PI) / 180;
        const x1 = 50 - Math.cos(radians) * 50;
        const y1 = 50 - Math.sin(radians) * 50;
        const x2 = 50 + Math.cos(radians) * 50;
        const y2 = 50 + Math.sin(radians) * 50;
        definitions.push(
          `<linearGradient id="${id}" x1="${x1}%" y1="${y1}%" x2="${x2}%" y2="${y2}%"><stop offset="0%" stop-color="${escapeXml(gradient.color1)}" stop-opacity="${gradient.opacity1 ?? 1}"/><stop offset="100%" stop-color="${escapeXml(gradient.color2)}" stop-opacity="${gradient.opacity2 ?? 1}"/></linearGradient>`,
        );
      }
    }

    if (element.dropShadow?.enabled) {
      const shadow = element.dropShadow;
      definitions.push(
        `<filter id="shadow-${escapeXml(element.id)}" x="-50%" y="-50%" width="200%" height="200%"><feDropShadow dx="${shadow.dx}" dy="${shadow.dy}" stdDeviation="${shadow.blur}" flood-color="${escapeXml(shadow.color)}" flood-opacity="${shadow.opacity}"/></filter>`,
      );
    }
  }

  const body = visibleElements
    .map((element) => {
      if (element.type === 'raw_svg') return serializeRawSvg(element);

      const fill = element.gradient?.enabled ? `url(#gradient-${element.id})` : element.fill;
      const paint = paintAttributes(element, fill);
      const transform = `transform="${transformFor(element)}"`;
      const filter = element.dropShadow?.enabled ? ` filter="url(#shadow-${escapeXml(element.id)})"` : '';
      const cx = element.x + element.width / 2;
      const cy = element.y + element.height / 2;

      switch (element.type) {
        case 'rect':
          return `<rect x="${element.x}" y="${element.y}" width="${element.width}" height="${element.height}" rx="${element.rx ?? 0}" ry="${element.ry ?? element.rx ?? 0}" ${paint} ${transform}${filter}/>`;
        case 'circle':
          return `<circle cx="${cx}" cy="${cy}" r="${Math.min(element.width, element.height) / 2}" ${paint} ${transform}${filter}/>`;
        case 'ellipse':
          return `<ellipse cx="${cx}" cy="${cy}" rx="${element.width / 2}" ry="${element.height / 2}" ${paint} ${transform}${filter}/>`;
        case 'line':
          return `<line x1="${element.x}" y1="${element.y}" x2="${element.x + element.width}" y2="${element.y + element.height}" ${paintAttributes(element, 'none')} ${transform}${filter}/>`;
        case 'arrow': {
          const markerId = `arrow-${escapeXml(element.id)}`;
          definitions.push(
            `<marker id="${markerId}" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto" markerUnits="strokeWidth"><path d="M0,0 L8,4 L0,8 z" fill="${escapeXml(element.stroke)}"/></marker>`,
          );
          return `<line x1="${element.x}" y1="${cy}" x2="${element.x + element.width}" y2="${cy}" ${paintAttributes(element, 'none')} marker-end="url(#${markerId})" ${transform}${filter}/>`;
        }
        case 'text':
          return `<text x="${element.x}" y="${element.y + (element.fontSize ?? 24)}" ${paint} font-size="${element.fontSize ?? 24}" font-family="${escapeXml(element.fontFamily ?? 'Inter')}" font-weight="${escapeXml(element.fontWeight ?? 'normal')}" ${transform}${filter}>${escapeXml(element.textContent)}</text>`;
        case 'polygon':
          return `<polygon points="${pointsForPolygon(element, Math.max(3, element.polygonSides ?? 6))}" ${paint} ${transform}${filter}/>`;
        case 'star':
          return `<polygon points="${pointsForStar(element)}" ${paint} ${transform}${filter}/>`;
        case 'path':
          return `<path d="${escapeXml(element.pathData)}" ${paint} transform="${transformFor(element)} ${pathTransformFor(element)}"${filter}/>`;
        default:
          return '';
      }
    })
    .filter(Boolean)
    .join('\n  ');

  const defs = definitions.length > 0 ? `\n  <defs>${definitions.join('')}</defs>` : '';
  const viewBox = config.viewBox;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${viewBox.x} ${viewBox.y} ${viewBox.w} ${viewBox.h}" width="${config.width}" height="${config.height}">${defs}\n  ${body}\n</svg>`;
};
