import type { ElectronAPI, ProjectFile, ProjectFolder, ProjectWorkbook } from './types/project';
import { computeOverallStatus } from '../shared/status';

function workbookNameForFolder(folderName: string, rawWorkbookName: string): string {
  const baseFolder = folderName.replace(/^PRJ_/i, '');
  const baseWorkbook = rawWorkbookName.replace(/\.xlsx$/i, '').trim();
  const joined = baseWorkbook.toLowerCase().startsWith(`${baseFolder}_`.toLowerCase())
    ? baseWorkbook
    : `${baseFolder}_${baseWorkbook}`;
  return joined.toLowerCase().endsWith('.xlsx') ? joined : `${joined}.xlsx`;
}

const statuses = [
  { value: 0, name: 'Non iniziato' },
  { value: 1, name: 'Work in progress' },
  { value: 2, name: 'Da verificare' },
  { value: 3, name: 'Concluso' },
  { value: 4, name: 'In Pausa' },
  { value: 5, name: 'Rimandato' },
  { value: 6, name: 'Da Definire' },
];

const folders: ProjectFolder[] = [
  {
    id: 'folder-1',
    name: '1Dash',
    path: '/SorgeProjects/1Dash',
    webUrl: '#',
    lastModifiedDateTime: new Date().toISOString(),
    files: [
      {
        id: 'demo-1',
        name: 'Piano lavori.xlsx',
        folderName: '1Dash',
        folderPath: '/SorgeProjects/1Dash',
        lastModifiedDateTime: new Date().toISOString(),
        webUrl: '#',
        size: 48220,
      },
      {
        id: 'demo-3',
        name: 'Checkpoint finale.xlsx',
        folderName: '1Dash',
        folderPath: '/SorgeProjects/1Dash',
        lastModifiedDateTime: new Date(Date.now() - 3600000).toISOString(),
        webUrl: '#',
        size: 32200,
      },
    ],
  },
  {
    id: 'folder-2',
    name: 'Template',
    path: '/SorgeProjects/Template',
    lastModifiedDateTime: new Date(Date.now() - 86400000).toISOString(),
    webUrl: '#',
    files: [
      {
        id: 'demo-2',
        name: 'template_piano_lavori_frontend.xlsx',
        folderName: 'Template',
        folderPath: '/SorgeProjects/Template',
        lastModifiedDateTime: new Date(Date.now() - 86400000).toISOString(),
        webUrl: '#',
        size: 21990,
      },
    ],
  },
];

function workbook(file: ProjectFile): ProjectWorkbook {
  return {
    file,
    sheetName: 'Piano lavori',
    format: 'piano-lavori',
    statuses,
    metadata: {
      links: [
        { label: 'Repository', url: 'https://github.com/sean-pegoraro-sorgetech/SorgeProjects' },
      ],
    },
    loadedAt: new Date().toISOString(),
    tasks: [
      {
        id: 'task-1',
        rowNumber: 5,
        area: 'Settings',
        task: 'Gestione permessi utente',
        owner: 'Sean',
        priority: 'Alta',
        dueDate: new Date(Date.now() + 86400000).toISOString().slice(0, 10),
        backendStatus: 'Da verificare',
        backendEstimateDays: 1.5,
        frontendStatus: 'Work in progress',
        frontendEstimateDays: 2,
        overallStatus: 'Work in progress',
        totalEstimateDays: 3.5,
        note1: '',
        note2: 'Test frontend in corso',
      },
      {
        id: 'task-2',
        rowNumber: 6,
        area: 'Calendario',
        task: 'Schedulazione con calendario organico',
        owner: 'Daniele',
        priority: 'Media',
        dueDate: '',
        backendStatus: 'Concluso',
        backendEstimateDays: 1,
        frontendStatus: 'Concluso',
        frontendEstimateDays: 2,
        overallStatus: 'Concluso',
        totalEstimateDays: 3,
        note1: '',
        note2: 'Da riverificare nel checkpoint finale',
      },
    ],
  };
}

export function installMockApi(): void {
  if (window.api) return;

  const api: ElectronAPI = {
    auth: {
      login: async () => ({ name: 'Sean Pegoraro', email: 'sean@sorgetech.it' }),
      logout: async () => undefined,
      getAccount: async () => ({ name: 'Sean Pegoraro', email: 'sean@sorgetech.it' }),
    },
    projects: {
      list: async () => folders,
      load: async (file) => workbook(file),
      save: async (project) => ({
        ...workbook(project.file),
        tasks: project.tasks.map((task) => ({
          ...task,
          overallStatus: computeOverallStatus(task.backendStatus, task.frontendStatus),
        })),
        statuses: project.statuses,
        sheetName: project.sheetName,
        format: project.format,
      }),
      create: async (input) => {
        const folder = await api.projects.createFolder({ name: input.name });
        return api.projects.createWorkbook({
          folderName: folder.name,
          folderPath: folder.path,
          fileName: 'Piano lavori',
          tasks: input.tasks,
        });
      },
      createFolder: async (input) => {
        const folder = {
          id: crypto.randomUUID(),
          name: input.name,
          path: `/SorgeProjects/${input.name}`,
          lastModifiedDateTime: new Date().toISOString(),
          webUrl: '#',
          files: [],
        };
        folders.unshift(folder);
        return folder;
      },
      createWorkbook: async (input) => {
        const name = workbookNameForFolder(input.folderName, input.fileName);
        const file = {
          id: crypto.randomUUID(),
          name,
          folderName: input.folderName,
          folderPath: input.folderPath,
          lastModifiedDateTime: new Date().toISOString(),
          webUrl: '#',
        };
        const folder = folders.find((item) => item.path === input.folderPath);
        folder?.files.unshift(file);
        return { ...workbook(file), tasks: input.tasks || [] };
      },
      deleteFolder: async (input) => {
        const index = folders.findIndex((folder) => folder.id === input.folder.id);
        if (index >= 0) folders.splice(index, 1);
      },
      deleteWorkbook: async (input) => {
        for (const folder of folders) {
          const index = folder.files.findIndex((file) => file.id === input.file.id);
          if (index >= 0) {
            folder.files.splice(index, 1);
            return;
          }
        }
      },
      archiveFolder: async (input) => {
        const index = folders.findIndex((folder) => folder.id === input.folder.id);
        if (index >= 0) folders.splice(index, 1);
      },
      listTemplates: async () => [
        { name: 'Piano lavori frontend', path: '/Template/template_piano_lavori_frontend.xlsx' },
      ],
    },
    settings: {
      get: async () => ({
        sharepoint: {
          mode: 'site',
          rootPath: '/Progetti',
          templatePath: '/Template/template_piano_lavori_frontend.xlsx',
          templateFolderPath: '/Template',
          archiveFolderName: '_Archivio',
        },
        defaults: {
          initialRows: 40,
          projectNamePrefix: '',
          owners: 'Sean\nDaniele',
          notificationEmail: '',
          teamsWebhookUrl: '',
        },
      }),
      set: async () => undefined,
    },
    azure: {
      get: async () => ({ clientId: '', tenantId: '' }),
      set: async () => undefined,
    },
    sharepointAdmin: {
      resolveDrive: async () => 'mock-drive-id',
    },
  };

  window.api = api;
}
