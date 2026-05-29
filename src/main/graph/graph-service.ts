import type { ProjectFile, ProjectFolder, ProjectTemplate, SharePointConfig } from '../../renderer/types/project';
import { getAccessToken } from '../auth/auth-service';

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

interface DriveItemResponse {
  id: string;
  name: string;
  webUrl?: string;
  size?: number;
  lastModifiedDateTime?: string;
  file?: unknown;
  folder?: unknown;
  parentReference?: {
    path?: string;
  };
}

function encodeDrivePath(value: string): string {
  return value
    .replace(/^\/+|\/+$/g, '')
    .split('/')
    .filter(Boolean)
    .map(encodeURIComponent)
    .join('/');
}

function getDriveBasePath(config: SharePointConfig): string {
  if (config.mode === 'group' && config.groupId) {
    return `/groups/${config.groupId}/drive`;
  }
  if (config.driveId) {
    return `/drives/${config.driveId}`;
  }
  throw new Error('Configurazione SharePoint incompleta. Inserire Group ID o Drive ID nelle Impostazioni.');
}

async function graphFetch(pathOrUrl: string, init: RequestInit = {}): Promise<Response> {
  const token = await getAccessToken();
  const url = pathOrUrl.startsWith('https://') ? pathOrUrl : `${GRAPH_ROOT}${pathOrUrl}`;
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(init.headers || {}),
    },
  });

  if (!response.ok) {
    const detail = await response.text().catch(() => '');
    if (response.status === 423 || detail.includes('resourceLocked')) {
      throw new Error(
        'Il file Excel e\' bloccato perche e\' aperto in Excel o nel browser. Chiudilo ovunque, attendi qualche secondo, poi riprova a salvare.'
      );
    }
    throw new Error(`Errore Microsoft Graph ${response.status}: ${detail || response.statusText}`);
  }

  return response;
}

async function graphJson<T>(path: string, init: RequestInit = {}): Promise<T> {
  const response = await graphFetch(path, {
    ...init,
    headers: {
      'Content-Type': 'application/json',
      ...(init.headers || {}),
    },
  });
  return response.json() as Promise<T>;
}

function toProjectFile(item: DriveItemResponse, folderName?: string, folderPath?: string): ProjectFile {
  return {
    id: item.id,
    name: item.name,
    folderName,
    folderPath,
    webUrl: item.webUrl,
    size: item.size,
    lastModifiedDateTime: item.lastModifiedDateTime,
  };
}

function toProjectFolder(item: DriveItemResponse, path: string, files: ProjectFile[]): ProjectFolder {
  return {
    id: item.id,
    name: item.name,
    path,
    webUrl: item.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
    files,
  };
}

async function listExcelFilesInPath(
  config: SharePointConfig,
  folderPath: string,
  folderName?: string
): Promise<ProjectFile[]> {
  const basePath = getDriveBasePath(config);
  const encodedFolderPath = encodeDrivePath(folderPath);
  const response = await graphJson<{ value: DriveItemResponse[] }>(
    `${basePath}/root:/${encodedFolderPath}:/children?$select=id,name,webUrl,lastModifiedDateTime,size,file&$orderby=lastModifiedDateTime desc`
  );

  return (response.value || [])
    .filter((item) => item.file && item.name.toLowerCase().endsWith('.xlsx') && !item.name.startsWith('~$'))
    .map((item) => toProjectFile(item, folderName, folderPath));
}

export async function listProjectFolders(config: SharePointConfig): Promise<ProjectFolder[]> {
  const basePath = getDriveBasePath(config);
  const rootPath = encodeDrivePath(config.rootPath || '/');
  const response = await graphJson<{ value: DriveItemResponse[] }>(
    `${basePath}/root:/${rootPath}:/children?$select=id,name,webUrl,lastModifiedDateTime,size,file,folder&$orderby=name`
  );

  const rootFiles = (response.value || [])
    .filter((item) => item.file && item.name.toLowerCase().endsWith('.xlsx') && !item.name.startsWith('~$'))
    .map((item) => toProjectFile(item, 'Root', config.rootPath || '/'));

  const archiveFolderName = (config.archiveFolderName || '_Archivio').toLowerCase();
  const folders = (response.value || []).filter((item) => item.folder && item.name.toLowerCase() !== archiveFolderName);
  const projectFolders = await Promise.all(
    folders.map(async (folder) => {
      const folderPath = `${config.rootPath || '/'}/${folder.name}`;
      const files = await listExcelFilesInPath(config, folderPath, folder.name);
      return toProjectFolder(folder, folderPath, files);
    })
  );

  const rootFolder: ProjectFolder | null = rootFiles.length
    ? {
        id: '__root__',
        name: 'Root',
        path: config.rootPath || '/',
        files: rootFiles,
        isRoot: true,
      }
    : null;

  return [
    ...(rootFolder ? [rootFolder] : []),
    ...projectFolders.sort((a, b) => a.name.localeCompare(b.name)),
  ];
}

export async function listWorkbookTemplates(config: SharePointConfig): Promise<ProjectTemplate[]> {
  const templateFolderPath = config.templateFolderPath?.trim();
  if (!templateFolderPath) return [];

  const files = await listExcelFilesInPath(config, templateFolderPath);
  return files.map((file) => ({
    name: file.name.replace(/\.xlsx$/i, ''),
    path: `${templateFolderPath.replace(/\/+$/g, '')}/${file.name}`,
  }));
}

export async function getProjectFileMetadata(
  config: SharePointConfig,
  itemId: string,
  folderName?: string,
  folderPath?: string
): Promise<ProjectFile> {
  const basePath = getDriveBasePath(config);
  const item = await graphJson<DriveItemResponse>(
    `${basePath}/items/${itemId}?$select=id,name,webUrl,lastModifiedDateTime,size,file`
  );
  return toProjectFile(item, folderName, folderPath);
}

export async function downloadProjectFile(config: SharePointConfig, itemId: string): Promise<Buffer> {
  const basePath = getDriveBasePath(config);
  const response = await graphFetch(`${basePath}/items/${itemId}/content`);
  return Buffer.from(await response.arrayBuffer());
}

export async function downloadFileByPath(config: SharePointConfig, relativePath: string): Promise<Buffer> {
  const basePath = getDriveBasePath(config);
  const encoded = encodeDrivePath(relativePath);
  const response = await graphFetch(`${basePath}/root:/${encoded}:/content`);
  return Buffer.from(await response.arrayBuffer());
}

export async function uploadProjectFileById(
  config: SharePointConfig,
  itemId: string,
  content: Buffer
): Promise<ProjectFile> {
  const basePath = getDriveBasePath(config);
  const response = await graphFetch(`${basePath}/items/${itemId}/content`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
    body: new Uint8Array(content) as BodyInit,
  });
  return toProjectFile((await response.json()) as DriveItemResponse);
}

export async function uploadProjectFileByPath(
  config: SharePointConfig,
  relativePath: string,
  content: Buffer
): Promise<ProjectFile> {
  const basePath = getDriveBasePath(config);
  const filePath = encodeDrivePath(relativePath);
  const response = await graphFetch(`${basePath}/root:/${filePath}:/content`, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    },
    body: new Uint8Array(content) as BodyInit,
  });
  return toProjectFile((await response.json()) as DriveItemResponse);
}

export async function deleteDriveItem(config: SharePointConfig, itemId: string): Promise<void> {
  const basePath = getDriveBasePath(config);
  await graphFetch(`${basePath}/items/${itemId}`, {
    method: 'DELETE',
  });
}

export async function createProjectFolder(config: SharePointConfig, folderName: string): Promise<ProjectFolder> {
  const basePath = getDriveBasePath(config);
  const rootPath = encodeDrivePath(config.rootPath || '/');
  const folder = await graphJson<DriveItemResponse>(`${basePath}/root:/${rootPath}:/children`, {
    method: 'POST',
    body: JSON.stringify({
      name: folderName,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'fail',
    }),
  });
  return toProjectFolder(folder, `${config.rootPath || '/'}/${folderName}`, []);
}

async function ensureArchiveFolder(config: SharePointConfig): Promise<ProjectFolder> {
  const basePath = getDriveBasePath(config);
  const rootPath = encodeDrivePath(config.rootPath || '/');
  const archiveFolderName = sanitizeFolderName(config.archiveFolderName || '_Archivio');
  const response = await graphJson<{ value: DriveItemResponse[] }>(
    `${basePath}/root:/${rootPath}:/children?$select=id,name,webUrl,lastModifiedDateTime,folder&$orderby=name`
  );

  const existing = (response.value || []).find(
    (item) => item.folder && item.name.toLowerCase() === archiveFolderName.toLowerCase()
  );
  if (existing) {
    return toProjectFolder(existing, `${config.rootPath || '/'}/${existing.name}`, []);
  }

  return createProjectFolder(config, archiveFolderName);
}

export async function archiveProjectFolder(config: SharePointConfig, folder: ProjectFolder): Promise<void> {
  const destination = await ensureArchiveFolder(config);
  const basePath = getDriveBasePath(config);
  await graphJson<DriveItemResponse>(`${basePath}/items/${folder.id}`, {
    method: 'PATCH',
    body: JSON.stringify({
      parentReference: {
        id: destination.id,
      },
    }),
  });
}

export async function resolveSiteDriveId(siteHostname: string, sitePath: string): Promise<string> {
  if (!siteHostname || !sitePath) {
    throw new Error('Site Hostname e Site Path sono obbligatori.');
  }

  const site = await graphJson<{ id: string }>(`/sites/${siteHostname}:/${sitePath}`);
  const drives = await graphJson<{ value: Array<{ id: string; name: string }> }>(`/sites/${site.id}/drives`);

  if (!drives.value || drives.value.length === 0) {
    throw new Error('Nessun drive trovato nel sito SharePoint.');
  }

  return drives.value[0].id;
}

export function sanitizeWorkbookName(name: string): string {
  const clean = name
    .replace(/[/\\:*?"<>|#%~&{}]/g, '_')
    .replace(/^\.+|\.+$/g, '')
    .trim();

  if (!clean) throw new Error('Nome progetto non valido.');
  return clean.toLowerCase().endsWith('.xlsx') ? clean : `${clean}.xlsx`;
}

export function sanitizeFolderName(name: string): string {
  const clean = name
    .replace(/[/\\:*?"<>|#%~&{}]/g, '_')
    .replace(/^\.+|\.+$/g, '')
    .trim();

  if (!clean) throw new Error('Nome progetto non valido.');
  return clean;
}
