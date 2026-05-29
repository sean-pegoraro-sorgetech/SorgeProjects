export interface SharePointConfig {
  mode: 'group' | 'site';
  groupId?: string;
  siteHostname?: string;
  sitePath?: string;
  driveId?: string;
  rootPath: string;
  templatePath?: string;
  templateFolderPath?: string;
  archiveFolderName?: string;
}

export interface AppSettings {
  sharepoint: SharePointConfig;
  defaults: {
    initialRows: number;
    projectNamePrefix: string;
    owners?: string;
    notificationEmail?: string;
    teamsWebhookUrl?: string;
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
  owner?: string;
  priority?: 'Bassa' | 'Media' | 'Alta' | 'Critica' | string;
  dueDate?: string;
  backendStatus: string;
  backendEstimateDays?: number | null;
  frontendStatus: string;
  frontendEstimateDays?: number | null;
  overallStatus: string;
  totalEstimateDays?: number | null;
  note1?: string;
  note2?: string;
}

export interface ProjectLink {
  label: string;
  url: string;
}

export interface ProjectMetadata {
  links: ProjectLink[];
}

export interface ProjectTemplate {
  name: string;
  path: string;
}

export interface ProjectWorkbook {
  file: ProjectFile;
  sheetName: string;
  format: WorkbookFormat;
  statuses: StatusOption[];
  tasks: ProjectTask[];
  metadata: ProjectMetadata;
  loadedAt: string;
}

export interface ProjectSaveInput {
  file: ProjectFile;
  sheetName: string;
  format: WorkbookFormat;
  statuses: StatusOption[];
  tasks: ProjectTask[];
  metadata?: ProjectMetadata;
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
  templatePath?: string;
  tasks?: ProjectTask[];
}

export interface DeleteFolderInput {
  folder: ProjectFolder;
}

export interface DeleteWorkbookInput {
  file: ProjectFile;
}

export interface ArchiveFolderInput {
  folder: ProjectFolder;
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
    archiveFolder(input: ArchiveFolderInput): Promise<void>;
    listTemplates(): Promise<ProjectTemplate[]>;
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
