export type ThemeName = 'dark' | 'light' | 'sorgetech' | 'forest' | 'graphite' | 'amber' | 'burgundy' | 'highContrast';
export type DensityName = 'comfortable' | 'compact';

export const THEMES: Array<{ value: ThemeName; label: string }> = [
  { value: 'dark', label: 'Scuro' },
  { value: 'light', label: 'Chiaro' },
  { value: 'sorgetech', label: 'Sorgetech' },
  { value: 'forest', label: 'Verde' },
  { value: 'graphite', label: 'Grafite' },
  { value: 'amber', label: 'Ambra' },
  { value: 'burgundy', label: 'Bordeaux' },
  { value: 'highContrast', label: 'Contrasto' },
];

export const DENSITIES: Array<{ value: DensityName; label: string; description: string }> = [
  { value: 'comfortable', label: 'Comoda', description: 'Spaziatura standard, migliore per uso quotidiano.' },
  { value: 'compact', label: 'Compatta', description: 'Righe e controlli piu stretti per vedere piu dati.' },
];

const THEME_KEY = 'project-step-manager:theme';
const DENSITY_KEY = 'project-step-manager:density';

function isThemeName(value: string | null): value is ThemeName {
  return THEMES.some((item) => item.value === value);
}

function isDensityName(value: string | null): value is DensityName {
  return value === 'comfortable' || value === 'compact';
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

export function getDensityPreference(): DensityName {
  const saved = window.localStorage.getItem(DENSITY_KEY);
  return isDensityName(saved) ? saved : 'comfortable';
}

export function applyDensity(density: DensityName): void {
  document.documentElement.dataset.density = density;
}

export function setDensityPreference(density: DensityName): void {
  window.localStorage.setItem(DENSITY_KEY, density);
  applyDensity(density);
}
