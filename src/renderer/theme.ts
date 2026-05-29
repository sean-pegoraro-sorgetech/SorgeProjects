export type ThemeName = 'dark' | 'light' | 'sorgetech' | 'forest';

export const THEMES: Array<{ value: ThemeName; label: string }> = [
  { value: 'dark', label: 'Scuro' },
  { value: 'light', label: 'Chiaro' },
  { value: 'sorgetech', label: 'Sorgetech' },
  { value: 'forest', label: 'Verde' },
];

const THEME_KEY = 'project-step-manager:theme';

function isThemeName(value: string | null): value is ThemeName {
  return value === 'dark' || value === 'light' || value === 'sorgetech' || value === 'forest';
}

export function getThemePreference(): ThemeName {
  const saved = window.localStorage.getItem(THEME_KEY);
  return isThemeName(saved) ? saved : 'dark';
}

export function applyTheme(theme: ThemeName): void {
  document.documentElement.dataset.theme = theme;
}

export function setThemePreference(theme: ThemeName): void {
  window.localStorage.setItem(THEME_KEY, theme);
  applyTheme(theme);
}
