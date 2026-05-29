export interface SharePointConfig {
  mode: 'group' | 'site';
  groupId?: string;
  siteHostname?: string;
  sitePath?: string;
  driveId?: string;
  rootPath: string;
  templatePath?: string;
}

export interface AppSettings {
  sharepoint: SharePointConfig;
  defaults: {
    initialRows: number;
    projectNamePrefix: string;
  };
}

export interface ProjectFile {
  id: string;
  name: string;
  folderName?: string;
  folderPath?: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  size?: number;
}

export interface ProjectFolder {
  id: string;
  name: string;
  path: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  files: ProjectFile[];
  isRoot?: boolean;
}

export type WorkbookFormat = 'piano-lavori' | 'legacy-status';

export interface StatusOption {
  value: number;
  name: string;
}

export interface ProjectTask {
  id: string;
  parentId?: string | null;
  level?: number;
  rowNumber?: number;
  phase?: string;
  area: string;
  task: string;
  backendStatus: string;
  backendEstimateDays?: number | null;
  frontendStatus: string;
  frontendEstimateDays?: number | null;
  overallStatus: string;
  totalEstimateDays?: number | null;
  note1?: string;
  note2?: string;
}

export interface ProjectWorkbook {
  file: ProjectFile;
  sheetName: string;
  format: WorkbookFormat;
  statuses: StatusOption[];
  tasks: ProjectTask[];
  loadedAt: string;
}

export interface ProjectSaveInput {
  file: ProjectFile;
  sheetName: string;
  format: WorkbookFormat;
  statuses: StatusOption[];
  tasks: ProjectTask[];
}

export interface NewProjectInput {
  name: string;
  tasks?: ProjectTask[];
}

export interface NewFolderInput {
  name: string;
}

export interface NewWorkbookInput {
  folderName: string;
  folderPath: string;
  fileName: string;
  tasks?: ProjectTask[];
}

export interface DeleteFolderInput {
  folder: ProjectFolder;
}

export interface DeleteWorkbookInput {
  file: ProjectFile;
}

export interface ElectronAPI {
  auth: {
    login(): Promise<{ name: string; email: string }>;
    logout(): Promise<void>;
    getAccount(): Promise<{ name: string; email: string } | null>;
  };
  projects: {
    list(): Promise<ProjectFolder[]>;
    load(file: ProjectFile): Promise<ProjectWorkbook>;
    save(project: ProjectSaveInput): Promise<ProjectWorkbook>;
    create(input: NewProjectInput): Promise<ProjectWorkbook>;
    createFolder(input: NewFolderInput): Promise<ProjectFolder>;
    createWorkbook(input: NewWorkbookInput): Promise<ProjectWorkbook>;
    deleteFolder(input: DeleteFolderInput): Promise<void>;
    deleteWorkbook(input: DeleteWorkbookInput): Promise<void>;
  };
  settings: {
    get(): Promise<AppSettings>;
    set(settings: Partial<AppSettings>): Promise<void>;
  };
  azure: {
    get(): Promise<{ clientId: string; tenantId: string }>;
    set(config: { clientId: string; tenantId: string }): Promise<void>;
  };
  sharepointAdmin: {
    resolveDrive(): Promise<string>;
  };
}

declare global {
  interface Window {
    api: ElectronAPI;
  }
}
