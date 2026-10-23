/**
 * Shared editor document model — single source of truth for element and
 * canvas construction. Used by the React editor (SvgEditor) and by the
 * headless EditorDriver, so both behave identically.
 */
import type { CanvasConfig, SvgElement, ToolType } from '../types/svg';
import { sanitizeSvg } from '../utils/svg';

export const defaultConfig = (): CanvasConfig => ({
  width: 24,
  height: 24,
  viewBox: { x: 0, y: 0, w: 24, h: 24 },
  backgroundColor: '#0f172a',
  gridVisible: true,
  gridSize: 1,
  snapToGrid: true,
});

export const defaultElement = (): SvgElement => ({
  id: 'rect-default',
  name: 'Icon Base 24x24',
  type: 'rect',
  x: 2,
  y: 2,
  width: 20,
  height: 20,
  rotation: 0,
  fill: '#3b82f6',
  fillOpacity: 0.85,
  stroke: '#1d4ed8',
  strokeWidth: 2,
  strokeOpacity: 1,
  rx: 4,
  ry: 4,
  visible: true,
  locked: false,
});

export const importedElement = (svgSource: string, name: string, x = 0, y = 0): SvgElement => {
  const safe = sanitizeSvg(svgSource);
  return {
    id: `imported-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    name,
    type: 'raw_svg',
    x,
    y,
    width: safe.width,
    height: safe.height,
    rotation: 0,
    fill: 'none',
    fillOpacity: 1,
    stroke: 'none',
    strokeWidth: 0,
    strokeOpacity: 1,
    rawSvgContent: safe.content,
    visible: true,
    locked: false,
  };
};

export const initialDocument = (source?: string): { config: CanvasConfig; elements: SvgElement[] } => {
  if (!source?.trim()) return { config: defaultConfig(), elements: [defaultElement()] };
  try {
    const element = importedElement(source, 'Imported SVG');
    return {
      config: {
        ...defaultConfig(),
        width: element.width,
        height: element.height,
        viewBox: { x: 0, y: 0, w: element.width, h: element.height },
        gridSize: element.width <= 64 ? 1 : Math.max(2, Math.round(element.width / 32)),
      },
      elements: [element],
    };
  } catch {
    return { config: defaultConfig(), elements: [defaultElement()] };
  }
};

/** Grid size used by the canvas size presets (same rule as the editor UI). */
export const presetGridSize = (width: number): number =>
  width <= 32 ? 1 : width <= 64 ? 2 : Math.max(4, Math.round(width / 32));

export type DrawToolType = Exclude<ToolType, 'select' | 'pan'>;

/**
 * Create a drawable shape with the editor's standard styling (identical for
 * the canvas drawing tools and the headless driver).
 */
export const createElement = (
  type: DrawToolType,
  x: number,
  y: number,
  sequence: number,
  id?: string,
): SvgElement => {
  const element: SvgElement = {
    id: id ?? `el-${Date.now()}-${sequence}`,
    name: `${type.toUpperCase()} ${sequence}`,
    type,
    x,
    y,
    width: 10,
    height: 10,
    rotation: 0,
    fill: type === 'line' || type === 'arrow' ? 'none' : '#3b82f6',
    fillOpacity: 1,
    stroke: '#1e40af',
    strokeWidth: 2,
    strokeOpacity: 1,
    polygonSides: 6,
    visible: true,
    locked: false,
  };
  if (type === 'text') {
    element.textContent = 'Sample Text';
    element.fontSize = 24;
    element.fontFamily = 'Inter';
    element.width = 140;
    element.height = 36;
  }
  if (type === 'path') element.pathViewBox = { x: 0, y: 0, w: element.width, h: element.height };
  return element;
};

export type AlignType = 'left' | 'centerH' | 'right' | 'top' | 'centerV' | 'bottom';

/** Alignment target position for an element inside a viewBox (editor rule). */
export const alignPosition = (
  element: Pick<SvgElement, 'x' | 'y' | 'width' | 'height'>,
  type: AlignType,
  viewBox: { x: number; y: number; w: number; h: number },
): { x: number; y: number } => {
  let x = element.x;
  let y = element.y;
  if (type === 'left') x = viewBox.x;
  if (type === 'centerH') x = viewBox.x + (viewBox.w - element.width) / 2;
  if (type === 'right') x = viewBox.x + viewBox.w - element.width;
  if (type === 'top') y = viewBox.y;
  if (type === 'centerV') y = viewBox.y + (viewBox.h - element.height) / 2;
  if (type === 'bottom') y = viewBox.y + viewBox.h - element.height;
  return { x, y };
};
