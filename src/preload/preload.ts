import { contextBridge, ipcRenderer } from 'electron';
import type {
  ElectronAPI,
  DeleteFolderInput,
  DeleteWorkbookInput,
  NewFolderInput,
  NewProjectInput,
  NewWorkbookInput,
  ProjectFile,
  ProjectSaveInput,
} from '../renderer/types/project';

const api: ElectronAPI = {
  auth: {
    login: () => ipcRenderer.invoke('auth:login'),
    logout: () => ipcRenderer.invoke('auth:logout'),
    getAccount: () => ipcRenderer.invoke('auth:getAccount'),
  },
  projects: {
    list: () => ipcRenderer.invoke('projects:list'),
    load: (file: ProjectFile) => ipcRenderer.invoke('projects:load', file),
    save: (project: ProjectSaveInput) => ipcRenderer.invoke('projects:save', project),
    create: (input: NewProjectInput) => ipcRenderer.invoke('projects:create', input),
    createFolder: (input: NewFolderInput) => ipcRenderer.invoke('projects:createFolder', input),
    createWorkbook: (input: NewWorkbookInput) => ipcRenderer.invoke('projects:createWorkbook', input),
    deleteFolder: (input: DeleteFolderInput) => ipcRenderer.invoke('projects:deleteFolder', input),
    deleteWorkbook: (input: DeleteWorkbookInput) => ipcRenderer.invoke('projects:deleteWorkbook', input),
  },
  settings: {
    get: () => ipcRenderer.invoke('settings:get'),
    set: (settings) => ipcRenderer.invoke('settings:set', settings),
  },
  azure: {
    get: () => ipcRenderer.invoke('azure:get'),
    set: (config: { clientId: string; tenantId: string }) => ipcRenderer.invoke('azure:set', config),
  },
  sharepointAdmin: {
    resolveDrive: () => ipcRenderer.invoke('sharepoint:resolveDrive'),
  },
};

contextBridge.exposeInMainWorld('api', api);
