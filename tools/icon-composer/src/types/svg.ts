export type ViewMode = 'editor' | 'viewer' | 'icon_hub';

export type ToolType =
  | 'select'
  | 'rect'
  | 'circle'
  | 'ellipse'
  | 'line'
  | 'arrow'
  | 'path'
  | 'text'
  | 'star'
  | 'polygon'
  | 'pan';

export interface CanvasConfig {
  width: number;
  height: number;
  viewBox: { x: number; y: number; w: number; h: number };
  backgroundColor: string;
  gridVisible: boolean;
  gridSize: number;
  snapToGrid: boolean;
}

export interface DropShadowFilter {
  enabled: boolean;
  dx: number;
  dy: number;
  blur: number;
  color: string;
  opacity: number;
}

export interface GradientDef {
  enabled: boolean;
  type: 'linear' | 'radial';
  angle: number; // 0 to 360 deg
  color1: string;
  color2: string;
  opacity1?: number;
  opacity2?: number;
}

export interface SvgElement {
  id: string;
  name: string;
  type: 'rect' | 'circle' | 'ellipse' | 'line' | 'arrow' | 'path' | 'text' | 'star' | 'polygon' | 'group' | 'raw_svg';
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipX?: boolean;
  flipY?: boolean;
  fill: string;
  fillOpacity: number;
  gradient?: GradientDef;
  stroke: string;
  strokeWidth: number;
  strokeOpacity: number;
  strokeDasharray?: string;
  strokeLinecap?: 'butt' | 'round' | 'square';
  strokeLinejoin?: 'miter' | 'round' | 'bevel';
  rx?: number;
  ry?: number;
  polygonSides?: number;
  textContent?: string;
  fontSize?: number;
  fontFamily?: string;
  fontWeight?: string;
  pathData?: string;
  /** Native path coordinates, mapped into the element's x/y/width/height slot. */
  pathViewBox?: { x: number; y: number; w: number; h: number };
  rawSvgContent?: string;
  dropShadow?: DropShadowFilter;
  visible: boolean;
  locked: boolean;
  children?: SvgElement[];
}

export interface HistoryState {
  elements: SvgElement[];
  description: string;
}

export interface LocalSvgFile {
  name: string;
  sizeBytes: number;
  lastModified: number;
  content: string;
  handle?: FileSystemFileHandle;
}

export interface FreeIcon {
  id: string;
  name: string;
  category: 'alignment' | 'charts' | 'tables' | 'shapes' | 'ui' | 'arrows' | 'files' | 'data' | 'devices' | 'media' | 'formatting' | 'tools' | 'general';
  tags: string[];
  viewBox?: string;
  svgContent: string;
  description?: string;
  license?: string;
  projectTarget?: string;
}
