export const STATUS_NAMES = {
  notStarted: 'Non iniziato',
  inProgress: 'Work in progress',
  review: 'Da verificare',
  done: 'Concluso',
  paused: 'In Pausa',
  deferred: 'Rimandato',
  undefined: 'Da Definire',
} as const;

export type KnownStatus = typeof STATUS_NAMES[keyof typeof STATUS_NAMES];

export function normalizeStatusName(value: string | undefined | null): KnownStatus {
  const text = (value || '').toLowerCase().replace(/\s+/g, ' ').trim();
  if (!text) return STATUS_NAMES.undefined;

  if (['0', 'non iniziato', 'not started', 'todo', 'to do'].includes(text)) return STATUS_NAMES.notStarted;
  if (['1', 'work in progress', 'wip', 'in progress', 'in corso', 'y wip'].includes(text)) {
    return STATUS_NAMES.inProgress;
  }
  if (['2', 'da verificare', 'vr', 'verifica', 'verify', 'review'].includes(text)) return STATUS_NAMES.review;
  if (['3', 'concluso', 'fatto', 'done', 'y', 'yes', 'si', 'true'].includes(text)) return STATUS_NAMES.done;
  if (['4', 'in pausa', 'pausa', 'paused', 'pause'].includes(text)) return STATUS_NAMES.paused;
  if (['5', 'rimandato', 'posticipato', 'deferred'].includes(text)) return STATUS_NAMES.deferred;
  if (['6', 'da definire', '-', 'na', 'n/a'].includes(text)) return STATUS_NAMES.undefined;

  return STATUS_NAMES.undefined;
}

export function computeOverallStatus(statusA: string, statusB: string): KnownStatus {
  const statuses = [normalizeStatusName(statusA), normalizeStatusName(statusB)];

  if (statuses.every((status) => status === STATUS_NAMES.notStarted)) return STATUS_NAMES.notStarted;
  if (statuses.includes(STATUS_NAMES.undefined)) return STATUS_NAMES.undefined;
  if (statuses.includes(STATUS_NAMES.paused)) return STATUS_NAMES.paused;
  if (statuses.includes(STATUS_NAMES.deferred)) return STATUS_NAMES.deferred;
  if (statuses.every((status) => status === STATUS_NAMES.done)) return STATUS_NAMES.done;
  if (statuses.includes(STATUS_NAMES.notStarted)) return STATUS_NAMES.inProgress;
  if (statuses.includes(STATUS_NAMES.inProgress)) return STATUS_NAMES.inProgress;
  if (statuses.includes(STATUS_NAMES.review)) return STATUS_NAMES.review;

  return STATUS_NAMES.undefined;
}

export function missingCounterpartLabel(backendStatus: string, frontendStatus: string): string | null {
  const backend = normalizeStatusName(backendStatus);
  const frontend = normalizeStatusName(frontendStatus);
  if (backend === STATUS_NAMES.notStarted && frontend !== STATUS_NAMES.notStarted) return 'Avvia backend';
  if (frontend === STATUS_NAMES.notStarted && backend !== STATUS_NAMES.notStarted) return 'Avvia frontend';
  return null;
}

export function statusClassName(status: string): string {
  const normalized = normalizeStatusName(status);
  if (normalized === STATUS_NAMES.done) return 'done';
  if (normalized === STATUS_NAMES.inProgress) return 'progress';
  if (normalized === STATUS_NAMES.review) return 'review';
  if (normalized === STATUS_NAMES.paused) return 'paused';
  if (normalized === STATUS_NAMES.deferred) return 'deferred';
  if (normalized === STATUS_NAMES.notStarted) return 'todo';
  return 'undefined';
}
