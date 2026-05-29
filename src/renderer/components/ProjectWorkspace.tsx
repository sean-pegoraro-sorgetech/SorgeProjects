import { type MouseEvent, useEffect, useMemo, useState } from 'react';
import {
  Archive,
  BarChart3,
  Bell,
  CalendarDays,
  ChevronDown,
  ChevronRight,
  Clipboard,
  Columns3,
  CornerDownRight,
  ExternalLink,
  FileSpreadsheet,
  FileText,
  Filter,
  Flag,
  Folder,
  LayoutDashboard,
  Link,
  ListPlus,
  Mail,
  PanelLeftOpen,
  Plus,
  RefreshCw,
  Save,
  Search,
  Send,
  Star,
  Table2,
  Trash2,
  Undo2,
  Users,
} from 'lucide-react';
import type {
  AppSettings,
  ProjectFile,
  ProjectFolder,
  ProjectLink,
  ProjectTask,
  ProjectTemplate,
  ProjectWorkbook,
  StatusOption,
} from '../types/project';
import { aggregateStatuses, computeOverallStatus, missingCounterpartLabel, statusClassName } from '../../shared/status';

const FAVORITES_KEY = 'project-step-manager:favorites';

interface Favorites {
  folderIds: string[];
  fileIds: string[];
}

type WorkspaceView = 'dashboard' | 'table' | 'kanban' | 'report';
type TaskFilter = 'all' | 'mine' | 'review' | 'paused' | 'overdue';

interface UserInfo {
  name: string;
  email: string;
}

interface WorkbookSummary {
  file: ProjectFile;
  folderName: string;
  total: number;
  done: number;
  review: number;
  paused: number;
  overdue: number;
  active: number;
  estimate: number;
  overall: string;
}

const emptyFavorites: Favorites = {
  folderIds: [],
  fileIds: [],
};

const PRIORITIES = ['Bassa', 'Media', 'Alta', 'Critica'];
const DEFAULT_LINKS = ['Repository', 'Test', 'Produzione', 'Documentazione', 'Teams', 'Ticket'];

function readFavorites(): Favorites {
  try {
    const raw = window.localStorage.getItem(FAVORITES_KEY);
    if (!raw) return emptyFavorites;
    const parsed = JSON.parse(raw) as Partial<Favorites>;
    return {
      folderIds: Array.isArray(parsed.folderIds) ? parsed.folderIds.filter((id): id is string => typeof id === 'string') : [],
      fileIds: Array.isArray(parsed.fileIds) ? parsed.fileIds.filter((id): id is string => typeof id === 'string') : [],
    };
  } catch {
    return emptyFavorites;
  }
}

function saveFavorites(favorites: Favorites) {
  window.localStorage.setItem(FAVORITES_KEY, JSON.stringify(favorites));
}

function toggleId(ids: string[], id: string): string[] {
  return ids.includes(id) ? ids.filter((item) => item !== id) : [id, ...ids];
}

function compareFavoriteName(aName: string, aFavorite: boolean, bName: string, bFavorite: boolean): number {
  if (aFavorite !== bFavorite) return aFavorite ? -1 : 1;
  return aName.localeCompare(bName, 'it', { sensitivity: 'base' });
}

function emptyTask(statuses: StatusOption[]): ProjectTask {
  const fallback = statuses.find((status) => status.name === 'Da Definire')?.name || statuses[0]?.name || 'Da Definire';
  return {
    id: `new-${crypto.randomUUID()}`,
    parentId: null,
    level: 0,
    area: '',
    task: '',
    owner: '',
    priority: 'Media',
    dueDate: '',
    backendStatus: fallback,
    backendEstimateDays: null,
    frontendStatus: fallback,
    frontendEstimateDays: null,
    overallStatus: fallback,
    totalEstimateDays: null,
    note1: '',
    note2: '',
  };
}

function formatDate(value?: string): string {
  if (!value) return '';
  return new Intl.DateTimeFormat('it-IT', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  }).format(new Date(value));
}

function formatDateOnly(value?: string): string {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat('it-IT', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }).format(date);
}

function numericValue(value: string): number | null {
  if (!value.trim()) return null;
  const parsed = Number.parseFloat(value.replace(',', '.'));
  return Number.isFinite(parsed) ? parsed : null;
}

function workbookLabel(file: ProjectFile): string {
  return file.folderName ? `${file.folderName} / ${file.name}` : file.name;
}

function isDone(task: ProjectTask): boolean {
  return statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)) === 'done';
}

function isOverdue(task: ProjectTask): boolean {
  if (!task.dueDate || isDone(task)) return false;
  const due = new Date(`${task.dueDate}T23:59:59`);
  return !Number.isNaN(due.getTime()) && due.getTime() < Date.now();
}

function summarizeWorkbook(workbook: ProjectWorkbook): WorkbookSummary {
  const tasks = rollupParentTasks(workbook.tasks);
  const rootTasks = tasks.filter((task) => !task.parentId);
  const done = tasks.filter(isDone).length;
  const review = tasks.filter((task) => statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)) === 'review').length;
  const paused = tasks.filter((task) => statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)) === 'paused').length;
  const overdue = tasks.filter(isOverdue).length;
  const active = tasks.filter((task) =>
    ['progress', 'review'].includes(statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)))
  ).length;
  const estimate = rootTasks.reduce((sum, task) => sum + (task.backendEstimateDays || 0) + (task.frontendEstimateDays || 0), 0);
  const overall = aggregateStatuses(rootTasks.map((task) => computeOverallStatus(task.backendStatus, task.frontendStatus)));
  return {
    file: workbook.file,
    folderName: workbook.file.folderName || 'Root',
    total: tasks.length,
    done,
    review,
    paused,
    overdue,
    active,
    estimate,
    overall,
  };
}

function mergeDefaultLinks(links: ProjectLink[]): ProjectLink[] {
  const existing = new Map(links.map((link) => [link.label, link.url]));
  const defaults = DEFAULT_LINKS.map((label) => ({ label, url: existing.get(label) || '' }));
  const custom = links.filter((link) => !DEFAULT_LINKS.includes(link.label));
  return [...defaults, ...custom];
}

function cloneWorkbook(workbook: ProjectWorkbook): ProjectWorkbook {
  return JSON.parse(JSON.stringify(workbook)) as ProjectWorkbook;
}

function ownerMatches(task: ProjectTask, user: UserInfo | null, myOwner?: string): boolean {
  if (!user || !task.owner) return false;
  const owner = task.owner.toLowerCase();
  const configured = (myOwner || '').trim().toLowerCase();
  if (configured && (owner === configured || owner.includes(configured))) return true;
  return owner.includes(user.email.toLowerCase()) || owner.includes((user.name || '').toLowerCase());
}

function buildChildrenMap(tasks: ProjectTask[]): Map<string, ProjectTask[]> {
  const childrenByParent = new Map<string, ProjectTask[]>();
  tasks.forEach((task) => {
    if (!task.parentId) return;
    const children = childrenByParent.get(task.parentId) || [];
    children.push(task);
    childrenByParent.set(task.parentId, children);
  });
  return childrenByParent;
}

function rollupTask(task: ProjectTask, children: ProjectTask[]): ProjectTask {
  if (children.length === 0) return task;

  const backendStatus = aggregateStatuses(children.map((child) => child.backendStatus));
  const frontendStatus = aggregateStatuses(children.map((child) => child.frontendStatus));
  const backendEstimateDays = children.reduce((sum, child) => sum + (child.backendEstimateDays || 0), 0) || null;
  const frontendEstimateDays = children.reduce((sum, child) => sum + (child.frontendEstimateDays || 0), 0) || null;

  return {
    ...task,
    backendStatus,
    frontendStatus,
    backendEstimateDays,
    frontendEstimateDays,
    overallStatus: computeOverallStatus(backendStatus, frontendStatus),
    totalEstimateDays: (backendEstimateDays || 0) + (frontendEstimateDays || 0) || null,
  };
}

function rollupParentTasks(tasks: ProjectTask[]): ProjectTask[] {
  const childrenByParent = buildChildrenMap(tasks);
  return tasks.map((task) => rollupTask(task, childrenByParent.get(task.id) || []));
}

export default function ProjectWorkspace() {
  const [folders, setFolders] = useState<ProjectFolder[]>([]);
  const [query, setQuery] = useState('');
  const [selectedFolderId, setSelectedFolderId] = useState<string | null>(null);
  const [selected, setSelected] = useState<ProjectWorkbook | null>(null);
  const [savedSnapshot, setSavedSnapshot] = useState<ProjectWorkbook | null>(null);
  const [loadingProjects, setLoadingProjects] = useState(false);
  const [loadingWorkbook, setLoadingWorkbook] = useState(false);
  const [saving, setSaving] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [newFolderName, setNewFolderName] = useState('');
  const [newWorkbookName, setNewWorkbookName] = useState('');
  const [libraryOpen, setLibraryOpen] = useState(true);
  const [favorites, setFavorites] = useState<Favorites>(() => readFavorites());
  const [expandedFolderIds, setExpandedFolderIds] = useState<string[]>([]);
  const [view, setView] = useState<WorkspaceView>('dashboard');
  const [taskFilter, setTaskFilter] = useState<TaskFilter>('all');
  const [ownerFilter, setOwnerFilter] = useState('');
  const [settings, setSettings] = useState<AppSettings | null>(null);
  const [currentUser, setCurrentUser] = useState<UserInfo | null>(null);
  const [templates, setTemplates] = useState<ProjectTemplate[]>([]);
  const [selectedTemplatePath, setSelectedTemplatePath] = useState('');
  const [dashboardSummaries, setDashboardSummaries] = useState<WorkbookSummary[]>([]);
  const [loadingDashboard, setLoadingDashboard] = useState(false);
  const [reportSent, setReportSent] = useState(false);

  const selectedFolder = useMemo(
    () => folders.find((folder) => folder.id === selectedFolderId) || null,
    [folders, selectedFolderId]
  );

  const upsertSummary = (workbook: ProjectWorkbook) => {
    const summary = summarizeWorkbook(workbook);
    setDashboardSummaries((current) => [
      summary,
      ...current.filter((item) => item.file.id !== workbook.file.id),
    ]);
  };

  const loadProjects = async () => {
    setLoadingProjects(true);
    setError(null);
    try {
      const items = await window.api.projects.list();
      setFolders(items);
      setSelectedFolderId((current) => {
        if (current && items.some((folder) => folder.id === current)) return current;
        return items[0]?.id || null;
      });
      setExpandedFolderIds((current) => {
        const existingIds = new Set(items.map((folder) => folder.id));
        const next = current.filter((id) => existingIds.has(id));
        if (next.length > 0) return next;
        return items[0]?.id ? [items[0].id] : [];
      });
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingProjects(false);
    }
  };

  useEffect(() => {
    loadProjects();
    window.api.settings.get().then(setSettings).catch(() => undefined);
    window.api.auth.getAccount().then(setCurrentUser).catch(() => undefined);
    window.api.projects.listTemplates().then(setTemplates).catch(() => setTemplates([]));
  }, []);

  useEffect(() => {
    if (!selected) setLibraryOpen(true);
  }, [selected]);

  const openProject = async (file: ProjectFile) => {
    if (dirty && !confirm('Ci sono modifiche non salvate. Vuoi cambiare Excel senza salvarle?')) return;
    setLoadingWorkbook(true);
    setError(null);
    try {
      const workbook = await window.api.projects.load(file);
      setSelected(cloneWorkbook(workbook));
      setSavedSnapshot(cloneWorkbook(workbook));
      upsertSummary(workbook);
      setDirty(false);
      setLibraryOpen(false);
      setView('table');
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingWorkbook(false);
    }
  };

  const createFolder = async () => {
    if (!newFolderName.trim()) return;
    setLoadingProjects(true);
    setError(null);
    try {
      const folder = await window.api.projects.createFolder({ name: newFolderName.trim() });
      setNewFolderName('');
      await loadProjects();
      setSelectedFolderId(folder.id);
      setExpandedFolderIds((current) => current.includes(folder.id) ? current : [folder.id, ...current]);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingProjects(false);
    }
  };

  const createWorkbook = async () => {
    if (!selectedFolder || !newWorkbookName.trim()) return;
    setLoadingWorkbook(true);
    setError(null);
    try {
      const workbook = await window.api.projects.createWorkbook({
        folderName: selectedFolder.name,
        folderPath: selectedFolder.path,
        fileName: newWorkbookName.trim(),
        templatePath: selectedTemplatePath || undefined,
      });
      setSelected(cloneWorkbook(workbook));
      setSavedSnapshot(cloneWorkbook(workbook));
      upsertSummary(workbook);
      setDirty(false);
      setLibraryOpen(false);
      setNewWorkbookName('');
      setExpandedFolderIds((current) => current.includes(selectedFolder.id) ? current : [selectedFolder.id, ...current]);
      await loadProjects();
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingWorkbook(false);
    }
  };

  const deleteFolder = async (folder: ProjectFolder) => {
    if (folder.isRoot) {
      setError('La cartella Root contiene file direttamente nella root SharePoint e non puo essere cancellata dall\'app.');
      return;
    }
    const count = folder.files.length;
    const message = count
      ? `Cancellare la cartella "${folder.name}" e tutti i ${count} Excel al suo interno?`
      : `Cancellare la cartella "${folder.name}"?`;
    if (!confirm(message)) return;

    setLoadingProjects(true);
    setError(null);
    try {
      await window.api.projects.deleteFolder({ folder });
      if (selected?.file.folderPath === folder.path) {
        setSelected(null);
        setSavedSnapshot(null);
        setDirty(false);
      }
      if (selectedFolderId === folder.id) setSelectedFolderId(null);
      setExpandedFolderIds((current) => current.filter((id) => id !== folder.id));
      await loadProjects();
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingProjects(false);
    }
  };

  const archiveFolder = async (folder: ProjectFolder) => {
    if (folder.isRoot) {
      setError('La cartella Root non puo essere archiviata dall\'app.');
      return;
    }
    if (!confirm(`Archiviare la cartella "${folder.name}"? Verrà spostata nella cartella archivio SharePoint.`)) return;

    setLoadingProjects(true);
    setError(null);
    try {
      await window.api.projects.archiveFolder({ folder });
      if (selected?.file.folderPath === folder.path) {
        setSelected(null);
        setSavedSnapshot(null);
        setDirty(false);
      }
      setDashboardSummaries((current) => current.filter((item) => item.file.folderPath !== folder.path));
      await loadProjects();
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingProjects(false);
    }
  };

  const deleteWorkbook = async (file: ProjectFile) => {
    if (!confirm(`Cancellare l'Excel "${file.name}"?`)) return;
    setLoadingProjects(true);
    setError(null);
    try {
      await window.api.projects.deleteWorkbook({ file });
      if (selected?.file.id === file.id) {
        setSelected(null);
        setSavedSnapshot(null);
        setDirty(false);
      }
      await loadProjects();
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingProjects(false);
    }
  };

  const saveProject = async () => {
    if (!selected) return;
    setSaving(true);
    setError(null);
    try {
      const tasksWithRollups = rollupParentTasks(selected.tasks);
      const children = buildChildrenMap(tasksWithRollups);
      const saved = await window.api.projects.save({
        file: selected.file,
        sheetName: selected.sheetName,
        format: selected.format,
        statuses: selected.statuses,
        metadata: selected.metadata,
        tasks: tasksWithRollups
          .filter((task) => task.task.trim() || (children.get(task.id)?.length || 0) > 0)
          .map((task) => ({
            ...task,
            overallStatus: computeOverallStatus(task.backendStatus, task.frontendStatus),
          })),
      });
      setSelected(cloneWorkbook(saved));
      setSavedSnapshot(cloneWorkbook(saved));
      upsertSummary(saved);
      setDirty(false);
      await loadProjects();
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const discardChanges = () => {
    if (!savedSnapshot) return;
    if (!confirm('Annullare tutte le modifiche non salvate e tornare all\'ultima versione caricata?')) return;
    setSelected(cloneWorkbook(savedSnapshot));
    setDirty(false);
    setError(null);
  };

  const updateTask = (id: string, patch: Partial<ProjectTask>) => {
    if (!selected) return;
    setSelected({
      ...selected,
      tasks: selected.tasks.map((task) => {
        if (task.id !== id) return task;
        const updated = { ...task, ...patch };
        return {
          ...updated,
          overallStatus: computeOverallStatus(updated.backendStatus, updated.frontendStatus),
        };
      }),
    });
    setDirty(true);
  };

  const addTask = () => {
    if (!selected) return;
    setSelected({ ...selected, tasks: [...selected.tasks, emptyTask(selected.statuses)] });
    setDirty(true);
  };

  const addSubtask = (parent: ProjectTask) => {
    if (!selected) return;
    const subtask = {
      ...emptyTask(selected.statuses),
      parentId: parent.id,
      level: (parent.level || 0) + 1,
      area: parent.area,
    };
    const parentIndex = selected.tasks.findIndex((task) => task.id === parent.id);
    const insertIndex = selected.tasks.reduce((lastIndex, task, index) => {
      if (task.id === parent.id || task.parentId === parent.id) return index + 1;
      return lastIndex;
    }, parentIndex + 1);
    const nextTasks = [...selected.tasks];
    nextTasks.splice(insertIndex, 0, subtask);
    setSelected({ ...selected, tasks: nextTasks });
    setDirty(true);
  };

  const deleteTask = (id: string) => {
    if (!selected) return;
    const children = selected.tasks.filter((task) => task.parentId === id);
    if (children.length > 0 && !confirm(`Eliminare anche ${children.length} sotto-task collegati?`)) return;
    setSelected({ ...selected, tasks: selected.tasks.filter((task) => task.id !== id && task.parentId !== id) });
    setDirty(true);
  };

  const toggleFolderFavorite = (event: MouseEvent<HTMLButtonElement>, folderId: string) => {
    event.stopPropagation();
    setFavorites((current) => {
      const next = { ...current, folderIds: toggleId(current.folderIds, folderId) };
      saveFavorites(next);
      return next;
    });
  };

  const toggleFileFavorite = (event: MouseEvent<HTMLButtonElement>, fileId: string) => {
    event.stopPropagation();
    setFavorites((current) => {
      const next = { ...current, fileIds: toggleId(current.fileIds, fileId) };
      saveFavorites(next);
      return next;
    });
  };

  const toggleFolderExpanded = (folder: ProjectFolder) => {
    setSelectedFolderId(folder.id);
    setExpandedFolderIds((current) =>
      current.includes(folder.id)
        ? current.filter((id) => id !== folder.id)
        : [folder.id, ...current]
    );
  };

  const refreshDashboard = async () => {
    setLoadingDashboard(true);
    setError(null);
    try {
      const summaries: WorkbookSummary[] = [];
      const files = folders.flatMap((folder) => folder.files);
      for (const file of files) {
        const workbook = await window.api.projects.load(file);
        summaries.push(summarizeWorkbook(workbook));
      }
      setDashboardSummaries(summaries);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingDashboard(false);
    }
  };

  const updateLink = (label: string, url: string) => {
    if (!selected) return;
    const links = mergeDefaultLinks(selected.metadata?.links || []).map((link) =>
      link.label === label ? { ...link, url } : link
    );
    setSelected({
      ...selected,
      metadata: { links },
    });
    setDirty(true);
  };

  const reportText = () => {
    const summaries = dashboardSummaries.length > 0
      ? dashboardSummaries
      : selected
        ? [summarizeWorkbook(selected)]
        : [];
    const totals = summaries.reduce(
      (acc, item) => ({
        projects: acc.projects + 1,
        tasks: acc.tasks + item.total,
        done: acc.done + item.done,
        review: acc.review + item.review,
        paused: acc.paused + item.paused,
        overdue: acc.overdue + item.overdue,
      }),
      { projects: 0, tasks: 0, done: 0, review: 0, paused: 0, overdue: 0 }
    );

    const lines = [
      `Report progetti tecnici - ${formatDateOnly(new Date().toISOString())}`,
      '',
      `Progetti: ${totals.projects}`,
      `Task: ${totals.tasks} (${totals.done} conclusi, ${totals.review} da verificare, ${totals.paused} in pausa, ${totals.overdue} in ritardo)`,
      '',
      ...summaries
        .sort((a, b) => b.overdue - a.overdue || b.review - a.review || a.folderName.localeCompare(b.folderName))
        .map((item) =>
          `- ${item.folderName} / ${item.file.name}: ${item.overall}, ${item.done}/${item.total} conclusi, ${item.review} da verificare, ${item.overdue} in ritardo`
        ),
    ];
    return lines.join('\n');
  };

  const copyReport = async () => {
    await navigator.clipboard.writeText(reportText());
    setReportSent(true);
    window.setTimeout(() => setReportSent(false), 2000);
  };

  const sendEmailReport = () => {
    const recipient = settings?.defaults.notificationEmail || '';
    const subject = encodeURIComponent('Report progetti tecnici');
    const body = encodeURIComponent(reportText());
    window.location.href = `mailto:${recipient}?subject=${subject}&body=${body}`;
  };

  const sendTeamsReport = async () => {
    const webhook = settings?.defaults.teamsWebhookUrl?.trim();
    if (!webhook) {
      setError('Webhook Teams non configurato nelle Impostazioni.');
      return;
    }
    setError(null);
    await fetch(webhook, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text: reportText().replace(/\n/g, '\n\n') }),
    });
    setReportSent(true);
    window.setTimeout(() => setReportSent(false), 2000);
  };

  const filteredFolders = useMemo(() => {
    const text = query.toLowerCase().trim();
    const visibleFolders = text
      ? folders
        .map((folder) => ({
          ...folder,
          files: folder.files.filter((file) =>
            `${folder.name} ${file.name}`.toLowerCase().includes(text)
          ),
        }))
        .filter((folder) => folder.name.toLowerCase().includes(text) || folder.files.length > 0)
      : folders;

    return visibleFolders
      .map((folder) => ({
        ...folder,
        files: [...folder.files].sort((a, b) =>
          compareFavoriteName(
            a.name,
            favorites.fileIds.includes(a.id),
            b.name,
            favorites.fileIds.includes(b.id)
          )
        ),
      }))
      .sort((a, b) =>
        compareFavoriteName(
          a.name,
          favorites.folderIds.includes(a.id),
          b.name,
          favorites.folderIds.includes(b.id)
        )
      );
  }, [favorites, folders, query]);

  const childrenByParent = useMemo(() => buildChildrenMap(selected?.tasks || []), [selected]);
  const rolledTasks = useMemo(() => rollupParentTasks(selected?.tasks || []), [selected]);
  const ownerOptions = useMemo(() => {
    const configured = (settings?.defaults.owners || '')
      .split(/[\n,;]/)
      .map((owner) => owner.trim())
      .filter(Boolean);
    const fromTasks = rolledTasks.map((task) => task.owner || '').filter(Boolean);
    const myOwner = settings?.defaults.myOwner ? [settings.defaults.myOwner] : [];
    return [...new Set([...myOwner, ...configured, ...fromTasks])].sort((a, b) => a.localeCompare(b, 'it'));
  }, [rolledTasks, settings]);

  const filteredTasks = useMemo(() => {
    return rolledTasks.filter((task) => {
      if (ownerFilter && task.owner !== ownerFilter) return false;
      if (taskFilter === 'mine' && !ownerMatches(task, currentUser, settings?.defaults.myOwner)) return false;
      if (taskFilter === 'review' && statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)) !== 'review') return false;
      if (taskFilter === 'paused' && statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)) !== 'paused') return false;
      if (taskFilter === 'overdue' && !isOverdue(task)) return false;
      return true;
    });
  }, [currentUser, ownerFilter, rolledTasks, settings, taskFilter]);

  const dashboardTotals = useMemo(() => {
    return dashboardSummaries.reduce(
      (acc, item) => ({
        projects: acc.projects + 1,
        tasks: acc.tasks + item.total,
        done: acc.done + item.done,
        review: acc.review + item.review,
        paused: acc.paused + item.paused,
        overdue: acc.overdue + item.overdue,
        estimate: acc.estimate + item.estimate,
      }),
      { projects: 0, tasks: 0, done: 0, review: 0, paused: 0, overdue: 0, estimate: 0 }
    );
  }, [dashboardSummaries]);

  const projectLinks = useMemo(() => mergeDefaultLinks(selected?.metadata?.links || []), [selected]);

  const summary = useMemo(() => {
    const tasks = filteredTasks;
    const estimateTasks = filteredTasks.filter((task) => !task.parentId);
    const done = tasks.filter((task) => statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)) === 'done').length;
    const active = tasks.filter((task) =>
      ['progress', 'review'].includes(statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)))
    ).length;
    const totalEstimate = estimateTasks.reduce((sum, task) => {
      const estimate = (task.backendEstimateDays || 0) + (task.frontendEstimateDays || 0);
      return sum + estimate;
    }, 0);
    return { total: tasks.length, done, active, totalEstimate };
  }, [filteredTasks]);

  return (
    <div className={`workspace-shell${selected && !libraryOpen ? ' library-hidden' : ''}`}>
      {libraryOpen && (
      <aside className="project-list">
        <div className="project-list-header">
          <div>
            <p className="eyebrow">SharePoint</p>
            <h1>Cartelle</h1>
          </div>
          <button className="icon-button" onClick={loadProjects} disabled={loadingProjects} title="Aggiorna">
            <RefreshCw size={18} className={loadingProjects ? 'spin-icon' : ''} />
          </button>
        </div>

        <div className="search-box">
          <Search size={16} />
          <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Cerca cartella o Excel" />
        </div>

        <div className="new-project">
          <input
            value={newFolderName}
            onChange={(event) => setNewFolderName(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') createFolder();
            }}
            placeholder="Nuova cartella progetto"
          />
          <button className="btn btn-primary btn-icon" onClick={createFolder} disabled={!newFolderName.trim()}>
            <Plus size={17} />
          </button>
        </div>

        <div className="project-items folder-tree">
          {filteredFolders.map((folder) => (
            <div key={folder.id} className={`folder-group${selectedFolderId === folder.id ? ' active' : ''}`}>
              <div className="folder-row">
                <button className="folder-main" onClick={() => toggleFolderExpanded(folder)}>
                  {expandedFolderIds.includes(folder.id) ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                  <Folder size={17} />
                  <span>
                    <strong>{folder.name}</strong>
                    <small>{folder.files.length} Excel</small>
                  </span>
                </button>
                <button
                  className={`row-action-btn favorite-btn${favorites.folderIds.includes(folder.id) ? ' active' : ''}`}
                  onClick={(event) => toggleFolderFavorite(event, folder.id)}
                  title={favorites.folderIds.includes(folder.id) ? 'Rimuovi dai preferiti' : 'Aggiungi cartella ai preferiti'}
                >
                  <Star size={15} />
                </button>
                {!folder.isRoot && (
                  <>
                    <button className="row-action-btn" onClick={() => archiveFolder(folder)} title="Archivia cartella progetto">
                      <Archive size={15} />
                    </button>
                    <button className="row-delete-btn" onClick={() => deleteFolder(folder)} title="Cancella cartella progetto">
                      <Trash2 size={15} />
                    </button>
                  </>
                )}
              </div>
              {expandedFolderIds.includes(folder.id) && (
                <div className="folder-files">
                  {folder.files.map((file) => (
                    <div
                      key={file.id}
                      className={`project-item file-row${selected?.file.id === file.id ? ' active' : ''}`}
                    >
                      <button className="file-main" onClick={() => openProject(file)}>
                        <FileSpreadsheet size={16} />
                        <span>
                          <strong>{file.name}</strong>
                          <small>{formatDate(file.lastModifiedDateTime)}</small>
                        </span>
                      </button>
                      <button
                        className={`row-action-btn favorite-btn${favorites.fileIds.includes(file.id) ? ' active' : ''}`}
                        onClick={(event) => toggleFileFavorite(event, file.id)}
                        title={favorites.fileIds.includes(file.id) ? 'Rimuovi dai preferiti' : 'Aggiungi Excel ai preferiti'}
                      >
                        <Star size={15} />
                      </button>
                      <button className="row-delete-btn" onClick={() => deleteWorkbook(file)} title="Cancella Excel">
                        <Trash2 size={15} />
                      </button>
                    </div>
                  ))}
                  {folder.files.length === 0 && <div className="empty-list compact">Nessun Excel in questa cartella.</div>}
                </div>
              )}
            </div>
          ))}
          {!loadingProjects && filteredFolders.length === 0 && (
            <div className="empty-list">Nessuna cartella trovata nella root configurata.</div>
          )}
        </div>
      </aside>
      )}

      <section className="project-detail">
        {error && <div className="alert error">{error}</div>}

        <div className="folder-toolbar">
          <div className="folder-toolbar-title">
            {selected && !libraryOpen && (
              <button className="icon-button" onClick={() => setLibraryOpen(true)} title="Mostra cartelle e Excel">
                <PanelLeftOpen size={18} />
              </button>
            )}
            <div>
              <p className="eyebrow">Cartella selezionata</p>
              <h2>{selectedFolder?.name || 'Nessuna cartella'}</h2>
            </div>
            {selectedFolder && (
              <button
                className={`icon-button favorite-btn${favorites.folderIds.includes(selectedFolder.id) ? ' active' : ''}`}
                onClick={(event) => toggleFolderFavorite(event, selectedFolder.id)}
                title={favorites.folderIds.includes(selectedFolder.id) ? 'Rimuovi cartella dai preferiti' : 'Aggiungi cartella ai preferiti'}
              >
                <Star size={17} />
              </button>
            )}
          </div>
          <div className="new-workbook">
            {templates.length > 0 && (
              <select
                value={selectedTemplatePath}
                onChange={(event) => setSelectedTemplatePath(event.target.value)}
                disabled={!selectedFolder}
                title="Template Excel"
              >
                <option value="">Template predefinito</option>
                {templates.map((template) => (
                  <option key={template.path} value={template.path}>{template.name}</option>
                ))}
              </select>
            )}
            <input
              value={newWorkbookName}
              onChange={(event) => setNewWorkbookName(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') createWorkbook();
              }}
              placeholder="Nuovo Excel nella cartella"
              disabled={!selectedFolder}
            />
            <button className="btn btn-primary" onClick={createWorkbook} disabled={!selectedFolder || !newWorkbookName.trim()}>
              <Plus size={16} /> Crea Excel
            </button>
          </div>
        </div>

        <div className="workspace-tabs">
          <button className={view === 'dashboard' ? 'active' : ''} onClick={() => setView('dashboard')}>
            <LayoutDashboard size={16} /> Dashboard
          </button>
          <button className={view === 'table' ? 'active' : ''} onClick={() => setView('table')} disabled={!selected}>
            <Table2 size={16} /> Tabella
          </button>
          <button className={view === 'kanban' ? 'active' : ''} onClick={() => setView('kanban')} disabled={!selected}>
            <Columns3 size={16} /> Kanban
          </button>
          <button className={view === 'report' ? 'active' : ''} onClick={() => setView('report')}>
            <FileText size={16} /> Report
          </button>
        </div>

        {view === 'dashboard' && (
          <div className="dashboard-view">
            <div className="dashboard-header">
              <div>
                <p className="eyebrow">Portfolio tecnico</p>
                <h2>Progetti in corso</h2>
              </div>
              <button className="btn btn-outline" onClick={refreshDashboard} disabled={loadingDashboard || folders.length === 0}>
                {loadingDashboard ? <span className="spinner" /> : <BarChart3 size={16} />}
                {loadingDashboard ? 'Analisi...' : 'Analizza Excel'}
              </button>
            </div>

            <div className="metrics-row portfolio">
              <div><span>{dashboardTotals.projects || folders.length}</span><small>Progetti/Excel</small></div>
              <div><span>{dashboardTotals.tasks || '-'}</span><small>Task tracciati</small></div>
              <div><span>{dashboardTotals.review || '-'}</span><small>Da verificare</small></div>
              <div><span>{dashboardTotals.overdue || '-'}</span><small>In ritardo</small></div>
            </div>

            <div className="portfolio-grid">
              {dashboardSummaries.map((item) => (
                <button key={item.file.id} className="portfolio-item" onClick={() => openProject(item.file)}>
                  <div>
                    <strong>{item.folderName}</strong>
                    <span>{item.file.name}</span>
                  </div>
                  <span className={`status-badge ${statusClassName(item.overall)}`}>{item.overall}</span>
                  <div className="portfolio-stats">
                    <small>{item.done}/{item.total} conclusi</small>
                    <small>{item.review} verifica</small>
                    <small>{item.overdue} ritardo</small>
                    <small>{item.estimate || '-'} gg</small>
                  </div>
                </button>
              ))}
              {!loadingDashboard && dashboardSummaries.length === 0 && (
                <div className="empty-list">Premi "Analizza Excel" per costruire la dashboard con task, ritardi e stati aggregati.</div>
              )}
            </div>
          </div>
        )}

        {view === 'report' && (
          <div className="report-view">
            <div className="dashboard-header">
              <div>
                <p className="eyebrow">Report settimanale</p>
                <h2>Riepilogo condivisibile</h2>
              </div>
              <div className="detail-actions">
                <button className="btn btn-outline" onClick={copyReport}><Clipboard size={16} /> Copia</button>
                <button className="btn btn-outline" onClick={sendEmailReport}><Mail size={16} /> Email</button>
                <button className="btn btn-outline" onClick={sendTeamsReport}><Send size={16} /> Teams</button>
              </div>
            </div>
            {reportSent && <div className="alert">Report inviato o copiato negli appunti.</div>}
            <pre className="report-box">{reportText()}</pre>
          </div>
        )}

        {!selected && !loadingWorkbook && view !== 'dashboard' && view !== 'report' && (
          <div className="empty-workspace">
            <FileSpreadsheet size={42} />
            <h2>Apri un Excel della cartella</h2>
            <p>Seleziona una cartella progetto, poi apri uno degli Excel al suo interno o creane uno nuovo.</p>
          </div>
        )}

        {loadingWorkbook && (
          <div className="empty-workspace">
            <div className="spinner large" />
            <p>Apertura workbook...</p>
          </div>
        )}

        {selected && !loadingWorkbook && view !== 'dashboard' && view !== 'report' && (
          <>
            <div className="detail-header">
              <div>
                <p className="eyebrow">{selected.format === 'piano-lavori' ? 'Piano lavori' : 'Formato storico'}</p>
                <h1>{workbookLabel(selected.file)}</h1>
                <p className="muted">Foglio: {selected.sheetName} - Ultimo caricamento: {formatDate(selected.loadedAt)}</p>
              </div>
              <div className="detail-actions">
                <button
                  className={`icon-button favorite-btn${favorites.fileIds.includes(selected.file.id) ? ' active' : ''}`}
                  onClick={(event) => toggleFileFavorite(event, selected.file.id)}
                  title={favorites.fileIds.includes(selected.file.id) ? 'Rimuovi Excel dai preferiti' : 'Aggiungi Excel ai preferiti'}
                >
                  <Star size={17} />
                </button>
                {selected.file.webUrl && (
                  <a className="btn btn-outline" href={selected.file.webUrl} target="_blank" rel="noreferrer">
                    <ExternalLink size={16} /> Excel
                  </a>
                )}
                {dirty && (
                  <>
                    <button className="btn btn-outline" onClick={discardChanges} disabled={saving}>
                      <Undo2 size={16} /> Annulla
                    </button>
                    <button className="btn btn-primary" onClick={saveProject} disabled={saving}>
                      {saving ? <span className="spinner" /> : <Save size={16} />}
                      {saving ? 'Salvataggio...' : 'Salva'}
                    </button>
                  </>
                )}
                {!dirty && <span className="sync-pill">Sincronizzato</span>}
              </div>
            </div>

            <div className="project-links-panel">
              {projectLinks.map((link) => (
                <label key={link.label}>
                  <span>{link.label}</span>
                  <div>
                    <Link size={14} />
                    <input
                      value={link.url}
                      onChange={(event) => updateLink(link.label, event.target.value)}
                      placeholder={`URL ${link.label.toLowerCase()}`}
                    />
                    {link.url && (
                      <a className="icon-button" href={link.url} target="_blank" rel="noreferrer" title={`Apri ${link.label}`}>
                        <ExternalLink size={15} />
                      </a>
                    )}
                  </div>
                </label>
              ))}
            </div>

            <div className="task-filter-bar">
              <div className="segmented-control">
                <button className={taskFilter === 'all' ? 'active' : ''} onClick={() => setTaskFilter('all')}>
                  <Filter size={15} /> Tutti
                </button>
                <button className={taskFilter === 'mine' ? 'active' : ''} onClick={() => setTaskFilter('mine')}>
                  <Users size={15} /> Miei
                </button>
                <button className={taskFilter === 'review' ? 'active' : ''} onClick={() => setTaskFilter('review')}>
                  <Bell size={15} /> Da verificare
                </button>
                <button className={taskFilter === 'paused' ? 'active' : ''} onClick={() => setTaskFilter('paused')}>
                  <Archive size={15} /> In pausa
                </button>
                <button className={taskFilter === 'overdue' ? 'active' : ''} onClick={() => setTaskFilter('overdue')}>
                  <CalendarDays size={15} /> In ritardo
                </button>
              </div>
              <label>
                <Users size={15} />
                <select value={ownerFilter} onChange={(event) => setOwnerFilter(event.target.value)}>
                  <option value="">Tutti gli owner</option>
                  {ownerOptions.map((owner) => <option key={owner} value={owner}>{owner}</option>)}
                </select>
              </label>
              {settings?.defaults.myOwner && <span className="owner-self-pill">Mio owner: {settings.defaults.myOwner}</span>}
            </div>

            <div className="metrics-row">
              <div><span>{summary.total}</span><small>Task</small></div>
              <div><span>{summary.done}</span><small>Conclusi</small></div>
              <div><span>{summary.active}</span><small>Attivi</small></div>
              <div><span>{summary.totalEstimate || '-'}</span><small>Giorni stimati</small></div>
            </div>

            {view === 'kanban' && (
              <div className="kanban-board">
                {selected.statuses.map((status) => {
                  const tasks = filteredTasks.filter((task) =>
                    computeOverallStatus(task.backendStatus, task.frontendStatus) === status.name
                  );
                  return (
                    <section key={status.name} className="kanban-column">
                      <div className="kanban-column-header">
                        <span className={`status-badge ${statusClassName(status.name)}`}>{status.name}</span>
                        <small>{tasks.length}</small>
                      </div>
                      <div className="kanban-cards">
                        {tasks.map((task) => (
                          <article key={task.id} className={`kanban-card ${task.parentId ? 'child' : ''}`}>
                            <div>
                              <strong>{task.task || 'Task senza nome'}</strong>
                              <small>{task.area || '-'}</small>
                            </div>
                            <div className="kanban-meta">
                              {task.owner && <span><Users size={13} /> {task.owner}</span>}
                              {task.priority && <span><Flag size={13} /> {task.priority}</span>}
                              {task.dueDate && <span className={isOverdue(task) ? 'overdue-text' : ''}><CalendarDays size={13} /> {formatDateOnly(task.dueDate)}</span>}
                            </div>
                          </article>
                        ))}
                        {tasks.length === 0 && <div className="empty-list compact">Nessun task.</div>}
                      </div>
                    </section>
                  );
                })}
              </div>
            )}

            {view === 'table' && (
            <>
            <div className="task-table-wrap">
              <datalist id="project-owner-options">
                {ownerOptions.map((owner) => <option key={owner} value={owner} />)}
              </datalist>
              <table className="task-table">
                <thead>
                  <tr>
                    <th>Ambito</th>
                    <th>Task</th>
                    <th>Owner</th>
                    <th>Priorità</th>
                    <th>Scadenza</th>
                    <th>Backend</th>
                    <th className="estimate-col">Stima (gg)</th>
                    <th>Frontend</th>
                    <th className="estimate-col">Stima (gg)</th>
                    <th>Stato</th>
                    <th>Nota 1</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {filteredTasks.map((task) => {
                    const childTasks = childrenByParent.get(task.id) || [];
                    const hasChildren = childTasks.length > 0;
                    const displayTask = rollupTask(task, childTasks);
                    const isSubtask = Boolean(task.parentId);
                    return (
                    <tr key={task.id} className={`${isSubtask ? 'subtask-row' : ''}${hasChildren ? ' parent-task-row' : ''}`}>
                      <td>
                        <input value={task.area} onChange={(event) => updateTask(task.id, { area: event.target.value })} />
                      </td>
                      <td className="task-name-cell">
                        <div className={`task-name-control${isSubtask ? ' child' : ''}`}>
                          {isSubtask && <CornerDownRight size={15} />}
                          <textarea
                            rows={1}
                            value={task.task}
                            onChange={(event) => updateTask(task.id, { task: event.target.value })}
                            placeholder={isSubtask ? 'Sotto-task' : 'Task'}
                          />
                        </div>
                        {hasChildren && <small className="computed-hint">Stato e stime calcolati dai sotto-task</small>}
                      </td>
                      <td>
                        <input
                          list="project-owner-options"
                          value={task.owner || ''}
                          onChange={(event) => updateTask(task.id, { owner: event.target.value })}
                          placeholder="Owner"
                        />
                      </td>
                      <td>
                        <select
                          value={task.priority || 'Media'}
                          className={`priority-select ${(task.priority || 'Media').toLowerCase()}`}
                          onChange={(event) => updateTask(task.id, { priority: event.target.value })}
                        >
                          {PRIORITIES.map((priority) => <option key={priority} value={priority}>{priority}</option>)}
                        </select>
                      </td>
                      <td>
                        <input
                          type="date"
                          value={task.dueDate || ''}
                          onChange={(event) => updateTask(task.id, { dueDate: event.target.value })}
                          className={isOverdue(task) ? 'overdue-input' : ''}
                        />
                      </td>
                      <td>
                        <select
                          value={displayTask.backendStatus}
                          className={`status-select ${statusClassName(displayTask.backendStatus)}`}
                          onChange={(event) => updateTask(task.id, { backendStatus: event.target.value })}
                          disabled={hasChildren}
                        >
                          {selected.statuses.map((status) => <option key={status.value} value={status.name}>{status.name}</option>)}
                        </select>
                      </td>
                      <td className="estimate-col">
                        <input
                          className="number-input"
                          value={displayTask.backendEstimateDays ?? ''}
                          onChange={(event) => updateTask(task.id, { backendEstimateDays: numericValue(event.target.value) })}
                          disabled={hasChildren}
                        />
                      </td>
                      <td>
                        <select
                          value={displayTask.frontendStatus}
                          className={`status-select ${statusClassName(displayTask.frontendStatus)}`}
                          onChange={(event) => updateTask(task.id, { frontendStatus: event.target.value })}
                          disabled={hasChildren}
                        >
                          {selected.statuses.map((status) => <option key={status.value} value={status.name}>{status.name}</option>)}
                        </select>
                      </td>
                      <td className="estimate-col">
                        <input
                          className="number-input"
                          value={displayTask.frontendEstimateDays ?? ''}
                          onChange={(event) => updateTask(task.id, { frontendEstimateDays: numericValue(event.target.value) })}
                          disabled={hasChildren}
                        />
                      </td>
                      <td>
                        <div className="computed-status-cell">
                          <span className={`status-badge ${statusClassName(computeOverallStatus(displayTask.backendStatus, displayTask.frontendStatus))}`}>
                            {computeOverallStatus(displayTask.backendStatus, displayTask.frontendStatus)}
                          </span>
                          {hasChildren && <small>Da {childTasks.length} sotto-task</small>}
                          {!hasChildren && missingCounterpartLabel(displayTask.backendStatus, displayTask.frontendStatus) && (
                            <small>{missingCounterpartLabel(displayTask.backendStatus, displayTask.frontendStatus)}</small>
                          )}
                        </div>
                      </td>
                      <td className="note-cell">
                        <textarea
                          rows={1}
                          value={task.note1 || ''}
                          onChange={(event) => updateTask(task.id, { note1: event.target.value })}
                          placeholder="Nota 1"
                        />
                      </td>
                      <td>
                        <div className="row-actions">
                          {!isSubtask && (
                            <button className="icon-button" onClick={() => addSubtask(task)} title="Aggiungi sotto-task">
                              <ListPlus size={16} />
                            </button>
                          )}
                          <button className="icon-button danger" onClick={() => deleteTask(task.id)} title="Elimina task">
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </td>
                    </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="table-footer">
              <button className="btn btn-outline" onClick={addTask}><Plus size={16} /> Aggiungi task</button>
            </div>
            </>
            )}
          </>
        )}
      </section>
    </div>
  );
}
