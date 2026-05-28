import { ipcMain } from 'electron';
import { initMsal, login, logout, getAccount } from './auth/auth-service';
import {
  createProjectFolder,
  deleteDriveItem,
  downloadFileByPath,
  downloadProjectFile,
  listProjectFolders,
  resolveSiteDriveId,
  sanitizeFolderName,
  sanitizeWorkbookName,
  uploadProjectFileById,
  uploadProjectFileByPath,
} from './graph/graph-service';
import {
  createProjectWorkbook,
  parseProjectWorkbook,
  writeProjectWorkbook,
} from './excel/workbook-service';
import { getAzureConfig, getSettings, setAzureConfig, setSettings } from './store/settings-store';
import type { NewProjectInput, ProjectFile, ProjectFolder, ProjectSaveInput } from '../renderer/types/project';

async function ensureMsalFromSettings(): Promise<void> {
  const azure = getAzureConfig();
  await initMsal(azure.clientId, azure.tenantId);
}

export function registerIpcHandlers(): void {
  ipcMain.handle('auth:login', async () => {
    await ensureMsalFromSettings();
    return login();
  });

  ipcMain.handle('auth:logout', async () => logout());

  ipcMain.handle('auth:getAccount', async () => {
    try {
      const azure = getAzureConfig();
      if (azure.clientId && azure.tenantId) await initMsal(azure.clientId, azure.tenantId);
      return getAccount();
    } catch {
      return null;
    }
  });

  ipcMain.handle('projects:list', async () => {
    const settings = getSettings();
    return listProjectFolders(settings.sharepoint);
  });

  ipcMain.handle('projects:load', async (_event, file: ProjectFile) => {
    const settings = getSettings();
    const content = await downloadProjectFile(settings.sharepoint, file.id);
    return parseProjectWorkbook(content, file);
  });

  ipcMain.handle('projects:save', async (_event, project: ProjectSaveInput) => {
    const settings = getSettings();
    const currentContent = await downloadProjectFile(settings.sharepoint, project.file.id);
    const updatedContent = await writeProjectWorkbook(currentContent, project, settings.defaults.initialRows);
    const updatedFile = await uploadProjectFileById(settings.sharepoint, project.file.id, updatedContent);
    return parseProjectWorkbook(updatedContent, updatedFile);
  });

  ipcMain.handle('projects:create', async (_event, input: NewProjectInput) => {
    const settings = getSettings();
    const folderName = sanitizeFolderName(`${settings.defaults.projectNamePrefix}${input.name}`);
    const filename = sanitizeWorkbookName('Piano lavori');
    const existing = await listProjectFolders(settings.sharepoint);
    if (existing.some((folder) => folder.name.toLowerCase() === folderName.toLowerCase())) {
      throw new Error(`Esiste gia una cartella progetto chiamata ${folderName}.`);
    }

    const templatePath = settings.sharepoint.templatePath?.trim();
    const templateContent = templatePath
      ? await downloadFileByPath(settings.sharepoint, templatePath)
      : undefined;
    const content = await createProjectWorkbook(input.name, input.tasks || [], settings.defaults.initialRows, templateContent);
    await createProjectFolder(settings.sharepoint, folderName);
    const file = await uploadProjectFileByPath(
      settings.sharepoint,
      `${settings.sharepoint.rootPath}/${folderName}/${filename}`,
      content
    );
    return parseProjectWorkbook(content, { ...file, folderName, folderPath: `${settings.sharepoint.rootPath}/${folderName}` });
  });

  ipcMain.handle('projects:createFolder', async (_event, input: { name: string }) => {
    const settings = getSettings();
    const folderName = sanitizeFolderName(`${settings.defaults.projectNamePrefix}${input.name}`);
    const existing = await listProjectFolders(settings.sharepoint);
    if (existing.some((folder) => folder.name.toLowerCase() === folderName.toLowerCase())) {
      throw new Error(`Esiste gia una cartella progetto chiamata ${folderName}.`);
    }
    return createProjectFolder(settings.sharepoint, folderName);
  });

  ipcMain.handle('projects:createWorkbook', async (_event, input: { folderName: string; folderPath: string; fileName: string; tasks?: any[] }) => {
    const settings = getSettings();
    const filename = sanitizeWorkbookName(input.fileName);
    const folders = await listProjectFolders(settings.sharepoint);
    const folder = folders.find((item) => item.path === input.folderPath);
    if (!folder) throw new Error(`Cartella progetto non trovata: ${input.folderName}`);
    if (folder.files.some((file) => file.name.toLowerCase() === filename.toLowerCase())) {
      throw new Error(`In ${input.folderName} esiste gia un Excel chiamato ${filename}.`);
    }

    const templatePath = settings.sharepoint.templatePath?.trim();
    const templateContent = templatePath
      ? await downloadFileByPath(settings.sharepoint, templatePath)
      : undefined;
    const content = await createProjectWorkbook(input.fileName, input.tasks || [], settings.defaults.initialRows, templateContent);
    const file = await uploadProjectFileByPath(settings.sharepoint, `${input.folderPath}/${filename}`, content);
    return parseProjectWorkbook(content, {
      ...file,
      folderName: input.folderName,
      folderPath: input.folderPath,
    });
  });

  ipcMain.handle('projects:deleteWorkbook', async (_event, input: { file: ProjectFile }) => {
    const settings = getSettings();
    await deleteDriveItem(settings.sharepoint, input.file.id);
  });

  ipcMain.handle('projects:deleteFolder', async (_event, input: { folder: ProjectFolder }) => {
    const settings = getSettings();
    if (input.folder.isRoot || input.folder.id === '__root__') {
      throw new Error('La cartella Root non puo essere cancellata dall\'app.');
    }
    await deleteDriveItem(settings.sharepoint, input.folder.id);
  });

  ipcMain.handle('settings:get', async () => getSettings());

  ipcMain.handle('settings:set', async (_event, partial) => {
    setSettings(partial);
  });

  ipcMain.handle('azure:get', async () => getAzureConfig());

  ipcMain.handle('azure:set', async (_event, config: { clientId: string; tenantId: string }) => {
    setAzureConfig(config);
  });

  ipcMain.handle('sharepoint:resolveDrive', async () => {
    const settings = getSettings();
    const sp = settings.sharepoint;

    if (sp.mode === 'group' && sp.groupId) {
      throw new Error('In modalita Group Drive non serve rilevare il Drive ID.');
    }

    if (sp.mode === 'site' && sp.siteHostname && sp.sitePath) {
      return resolveSiteDriveId(sp.siteHostname, sp.sitePath);
    }

    throw new Error('Configurazione SharePoint incompleta. Compilare Site Hostname e Site Path.');
  });
}
