import Store from 'electron-store';
import type { AppSettings } from '../../renderer/types/project';

const defaults: AppSettings = {
  sharepoint: {
    mode: 'site',
    siteHostname: 'sorgetech.sharepoint.com',
    sitePath: '/sites/Clienti-Documentazionetecnicainterna',
    rootPath: '/SorgeProjects',
    templatePath: '',
    templateFolderPath: '',
    archiveFolderName: '_Archivio',
  },
  defaults: {
    initialRows: 40,
    projectNamePrefix: '',
    owners: '',
    notificationEmail: '',
    teamsWebhookUrl: '',
  },
};

interface StoreSchema {
  settings: AppSettings;
  azure: {
    clientId: string;
    tenantId: string;
  };
}

const store = new Store<StoreSchema>({
  defaults: {
    settings: defaults,
    azure: {
      clientId: '',
      tenantId: '',
    },
  },
});

export function getSettings(): AppSettings {
  return store.get('settings');
}

export function setSettings(partial: Partial<AppSettings>): void {
  const current = store.get('settings');
  store.set('settings', {
    ...current,
    ...partial,
    sharepoint: partial.sharepoint
      ? { ...current.sharepoint, ...partial.sharepoint }
      : current.sharepoint,
    defaults: partial.defaults
      ? { ...current.defaults, ...partial.defaults }
      : current.defaults,
  });
}

export function getAzureConfig(): { clientId: string; tenantId: string } {
  return store.get('azure');
}

export function setAzureConfig(config: { clientId: string; tenantId: string }): void {
  store.set('azure', config);
}

export { store };
