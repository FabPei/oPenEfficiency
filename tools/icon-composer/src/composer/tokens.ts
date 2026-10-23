import type { AccentRole, ActionSemantic, TokenTheme } from './types';

/**
 * Semantic accent tokens mirroring the oPenEfficiency theme resources
 * (UI/Styles/Themes/Light.xaml + Dark.xaml). Values are design tokens resolved
 * centrally here – never hardcoded into individual SVG paths.
 */
export const OFFICE_TOKENS: Record<TokenTheme, Record<AccentRole | 'Foreground', string>> = {
  light: {
    Foreground: '#424242',
    ActionBlue: '#2563EB',
    CreateGreen: '#059669',
    SuccessGreen: '#059669',
    AttentionYellow: '#D97706',
    DestructiveRed: '#DC2626',
    ContentColor: '#2563EB',
  },
  dark: {
    Foreground: '#D6D6D6',
    ActionBlue: '#3B82F6',
    CreateGreen: '#10B981',
    SuccessGreen: '#10B981',
    AttentionYellow: '#F59E0B',
    DestructiveRed: '#EF4444',
    ContentColor: '#3B82F6',
  },
};

/** Action type -> semantic accent role (consumed by the oPenEfficiency icon parser). */
export const ACTION_TYPE_ROLE: Record<ActionSemantic, AccentRole> = {
  add: 'CreateGreen',
  confirm: 'SuccessGreen',
  move: 'ActionBlue',
  refresh: 'ActionBlue',
  delete: 'DestructiveRed',
  warning: 'AttentionYellow',
  custom: 'ContentColor',
};

export const ACTION_TYPE_LABEL: Record<ActionSemantic, string> = {
  add: 'Add / Plus',
  confirm: 'Confirm / Check',
  move: 'Move / Forward',
  refresh: 'Refresh / Sync',
  delete: 'Delete / Remove',
  warning: 'Warning / Attention',
  custom: 'Custom',
};

export const resolveAccentColor = (role: AccentRole, theme: TokenTheme): string =>
  OFFICE_TOKENS[theme][role];

export const resolveBaseColor = (theme: TokenTheme): string => OFFICE_TOKENS[theme].Foreground;
