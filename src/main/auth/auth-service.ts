import {
  AccountInfo,
  AuthenticationResult,
  Configuration,
  LogLevel,
  PublicClientApplication,
} from '@azure/msal-node';
import {
  DataProtectionScope,
  PersistenceCachePlugin,
  PersistenceCreator,
} from '@azure/msal-node-extensions';
import { app } from 'electron';
import path from 'path';

const SCOPES = ['User.Read', 'Files.ReadWrite.All', 'Sites.Read.All'];

let msalInstance: PublicClientApplication | null = null;
let currentAccount: AccountInfo | null = null;

async function createCachePlugin() {
  const cachePath = path.join(app.getPath('userData'), 'msal-cache.json');

  const persistence = await PersistenceCreator.createPersistence({
    cachePath,
    dataProtectionScope: DataProtectionScope.CurrentUser,
    serviceName: 'ProjectStepManager',
    accountName: 'MSALCache',
    usePlaintextFileOnLinux: false,
  });

  return new PersistenceCachePlugin(persistence);
}

export async function initMsal(clientId?: string, tenantId?: string): Promise<void> {
  if (!clientId || !tenantId) {
    throw new Error('Azure App Registration non configurata. Inserire Client ID e Tenant ID nelle Impostazioni.');
  }

  const cachePlugin = await createCachePlugin();
  const config: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: { cachePlugin },
    system: {
      loggerOptions: {
        loggerCallback: (_level, message) => {
          if (process.env.NODE_ENV === 'development') console.log('[MSAL]', message);
        },
        logLevel: LogLevel.Warning,
        piiLoggingEnabled: false,
      },
    },
  };

  msalInstance = new PublicClientApplication(config);
}

export async function login(): Promise<{ name: string; email: string }> {
  if (!msalInstance) throw new Error('MSAL non inizializzato');

  const result: AuthenticationResult = await msalInstance.acquireTokenInteractive({
    scopes: SCOPES,
    openBrowser: async (url) => {
      const { shell } = await import('electron');
      await shell.openExternal(url);
    },
    successTemplate: '<h1>Login riuscito. Puoi chiudere questa finestra.</h1>',
    errorTemplate: '<h1>Errore durante il login. Riprova.</h1>',
  });

  currentAccount = result.account;
  return {
    name: result.account?.name || '',
    email: result.account?.username || '',
  };
}

export async function logout(): Promise<void> {
  if (!msalInstance) return;

  const cache = msalInstance.getTokenCache();
  const accounts = await cache.getAllAccounts();
  for (const account of accounts) {
    await cache.removeAccount(account);
  }
  currentAccount = null;
}

export async function getAccount(): Promise<{ name: string; email: string } | null> {
  if (!msalInstance) return null;

  const cache = msalInstance.getTokenCache();
  const accounts = await cache.getAllAccounts();
  if (accounts.length === 0) return null;

  currentAccount = accounts[0];
  return {
    name: currentAccount.name || '',
    email: currentAccount.username || '',
  };
}

export async function getAccessToken(): Promise<string> {
  if (!msalInstance) throw new Error('MSAL non inizializzato');

  if (!currentAccount) {
    const cache = msalInstance.getTokenCache();
    const accounts = await cache.getAllAccounts();
    if (accounts.length === 0) throw new Error('Nessun account. Effettuare il login.');
    currentAccount = accounts[0];
  }

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: SCOPES,
      account: currentAccount,
    });
    return result.accessToken;
  } catch {
    const result = await msalInstance.acquireTokenInteractive({
      scopes: SCOPES,
      openBrowser: async (url) => {
        const { shell } = await import('electron');
        await shell.openExternal(url);
      },
    });
    currentAccount = result.account;
    return result.accessToken;
  }
}
