/** 3×3 anchor grid used by both the visual composer and programmatic AI API. */
export const ACTION_POSITIONS = [
  'top-left', 'top-center', 'top-right',
  'middle-left', 'center', 'middle-right',
  'bottom-left', 'bottom-center', 'bottom-right',
] as const;
export type ActionPosition = (typeof ACTION_POSITIONS)[number];

export type ActionSemantic =
  | 'add'
  | 'confirm'
  | 'move'
  | 'refresh'
  | 'delete'
  | 'warning'
  | 'custom';

export type AccentRole =
  | 'ActionBlue'
  | 'CreateGreen'
  | 'SuccessGreen'
  | 'AttentionYellow'
  | 'DestructiveRed'
  | 'ContentColor';

export type TokenTheme = 'light' | 'dark';

export interface ComposerParams {
  actionType: ActionSemantic;
  position: ActionPosition;
  /** Action visual size as a fraction of the main icon visual size (0.2 – 1.5). */
  actionSize: number;
  /** Transparent gap between main icon and action icon, in 24-unit target space. */
  knockoutPadding: number;
  autoColor: boolean;
  accentRole: AccentRole;
  /** Resolved accent color (hex). Used when autoColor is on (from tokens) or as manual color. */
  accentColor: string;
  /** Main icon ink color (hex). Applied via the `color` property, sources keep currentColor. */
  baseColor: string;
  tokenTheme: TokenTheme;
}

export interface ComposerConfig {
  mainSvg: string;
  actionSvg: string;
  params: ComposerParams;
  /** Id namespace of an existing composition (restored on re-import). */
  uid?: string;
}

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ComposerLayer {
  name: string;
  svg: string;
}

export interface ComposedIcon {
  /** Single clean SVG document containing main-icon, action-knockout and action-icon. */
  svg: string;
  /** Same visuals without composer metadata (for flat export). */
  flattenedSvg: string;
  /** Editor layers: main-icon (incl. knockout mask) and action-icon. */
  layers: ComposerLayer[];
  /** True when the composition embeds <metadata> for later parametric re-editing. */
  hasMetadata: boolean;
  mainBounds: Rect;
  actionBounds: Rect;
  /** Points inside the transparent gap (target space) — QA/visual tests only. */
  probes: Array<{ x: number; y: number }>;
  warnings: string[];
}

export class ComposerError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ComposerError';
  }
}
