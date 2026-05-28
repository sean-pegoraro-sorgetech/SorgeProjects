import { useEffect, useMemo, useState } from 'react';
import {
  ExternalLink,
  FileSpreadsheet,
  Folder,
  Plus,
  RefreshCw,
  Save,
  Search,
  Trash2,
} from 'lucide-react';
import type { ProjectFile, ProjectFolder, ProjectTask, ProjectWorkbook, StatusOption } from '../types/project';
import { computeOverallStatus, missingCounterpartLabel, statusClassName } from '../../shared/status';

function emptyTask(statuses: StatusOption[]): ProjectTask {
  const fallback = statuses.find((status) => status.name === 'Da Definire')?.name || statuses[0]?.name || 'Da Definire';
  return {
    id: `new-${crypto.randomUUID()}`,
    area: '',
    task: '',
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

function numericValue(value: string): number | null {
  if (!value.trim()) return null;
  const parsed = Number.parseFloat(value.replace(',', '.'));
  return Number.isFinite(parsed) ? parsed : null;
}

function workbookLabel(file: ProjectFile): string {
  return file.folderName ? `${file.folderName} / ${file.name}` : file.name;
}

export default function ProjectWorkspace() {
  const [folders, setFolders] = useState<ProjectFolder[]>([]);
  const [query, setQuery] = useState('');
  const [selectedFolderId, setSelectedFolderId] = useState<string | null>(null);
  const [selected, setSelected] = useState<ProjectWorkbook | null>(null);
  const [loadingProjects, setLoadingProjects] = useState(false);
  const [loadingWorkbook, setLoadingWorkbook] = useState(false);
  const [saving, setSaving] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [newFolderName, setNewFolderName] = useState('');
  const [newWorkbookName, setNewWorkbookName] = useState('');

  const selectedFolder = useMemo(
    () => folders.find((folder) => folder.id === selectedFolderId) || null,
    [folders, selectedFolderId]
  );

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
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoadingProjects(false);
    }
  };

  useEffect(() => {
    loadProjects();
  }, []);

  const openProject = async (file: ProjectFile) => {
    if (dirty && !confirm('Ci sono modifiche non salvate. Vuoi cambiare Excel senza salvarle?')) return;
    setLoadingWorkbook(true);
    setError(null);
    try {
      const workbook = await window.api.projects.load(file);
      setSelected(workbook);
      setDirty(false);
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
      });
      setSelected(workbook);
      setDirty(false);
      setNewWorkbookName('');
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
        setDirty(false);
      }
      if (selectedFolderId === folder.id) setSelectedFolderId(null);
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
      const saved = await window.api.projects.save({
        file: selected.file,
        sheetName: selected.sheetName,
        format: selected.format,
        statuses: selected.statuses,
        tasks: selected.tasks
          .filter((task) => task.task.trim())
          .map((task) => ({
            ...task,
            overallStatus: computeOverallStatus(task.backendStatus, task.frontendStatus),
          })),
      });
      setSelected(saved);
      setDirty(false);
      await loadProjects();
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
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

  const deleteTask = (id: string) => {
    if (!selected) return;
    setSelected({ ...selected, tasks: selected.tasks.filter((task) => task.id !== id) });
    setDirty(true);
  };

  const filteredFolders = useMemo(() => {
    const text = query.toLowerCase().trim();
    if (!text) return folders;
    return folders
      .map((folder) => ({
        ...folder,
        files: folder.files.filter((file) =>
          `${folder.name} ${file.name}`.toLowerCase().includes(text)
        ),
      }))
      .filter((folder) => folder.name.toLowerCase().includes(text) || folder.files.length > 0);
  }, [folders, query]);

  const summary = useMemo(() => {
    const tasks = selected?.tasks || [];
    const done = tasks.filter((task) => statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)) === 'done').length;
    const active = tasks.filter((task) =>
      ['progress', 'review'].includes(statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus)))
    ).length;
    const totalEstimate = tasks.reduce((sum, task) => {
      const estimate = (task.backendEstimateDays || 0) + (task.frontendEstimateDays || 0);
      return sum + estimate;
    }, 0);
    return { total: tasks.length, done, active, totalEstimate };
  }, [selected]);

  return (
    <div className="workspace-shell">
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
                <button className="folder-main" onClick={() => setSelectedFolderId(folder.id)}>
                  <Folder size={17} />
                  <span>
                    <strong>{folder.name}</strong>
                    <small>{folder.files.length} Excel</small>
                  </span>
                </button>
                {!folder.isRoot && (
                  <button className="row-delete-btn" onClick={() => deleteFolder(folder)} title="Cancella cartella progetto">
                    <Trash2 size={15} />
                  </button>
                )}
              </div>
              {selectedFolderId === folder.id && (
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

      <section className="project-detail">
        {error && <div className="alert error">{error}</div>}

        <div className="folder-toolbar">
          <div>
            <p className="eyebrow">Cartella selezionata</p>
            <h2>{selectedFolder?.name || 'Nessuna cartella'}</h2>
          </div>
          <div className="new-workbook">
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

        {!selected && !loadingWorkbook && (
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

        {selected && !loadingWorkbook && (
          <>
            <div className="detail-header">
              <div>
                <p className="eyebrow">{selected.format === 'piano-lavori' ? 'Piano lavori' : 'Formato storico'}</p>
                <h1>{workbookLabel(selected.file)}</h1>
                <p className="muted">Foglio: {selected.sheetName} - Ultimo caricamento: {formatDate(selected.loadedAt)}</p>
              </div>
              <div className="detail-actions">
                {selected.file.webUrl && (
                  <a className="btn btn-outline" href={selected.file.webUrl} target="_blank" rel="noreferrer">
                    <ExternalLink size={16} /> Excel
                  </a>
                )}
                {dirty && (
                  <button className="btn btn-primary" onClick={saveProject} disabled={saving}>
                    {saving ? <span className="spinner" /> : <Save size={16} />}
                    {saving ? 'Salvataggio...' : 'Salva'}
                  </button>
                )}
                {!dirty && <span className="sync-pill">Sincronizzato</span>}
              </div>
            </div>

            <div className="metrics-row">
              <div><span>{summary.total}</span><small>Task</small></div>
              <div><span>{summary.done}</span><small>Conclusi</small></div>
              <div><span>{summary.active}</span><small>Attivi</small></div>
              <div><span>{summary.totalEstimate || '-'}</span><small>Giorni stimati</small></div>
            </div>

            <div className="task-table-wrap">
              <table className="task-table">
                <thead>
                  <tr>
                    <th>Ambito</th>
                    <th>Task</th>
                    <th>Backend</th>
                    <th>Stima</th>
                    <th>Frontend</th>
                    <th>Stima</th>
                    <th>Stato</th>
                    <th>Nota 1</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {selected.tasks.map((task) => (
                    <tr key={task.id}>
                      <td>
                        <input value={task.area} onChange={(event) => updateTask(task.id, { area: event.target.value })} />
                      </td>
                      <td className="task-name-cell">
                        <textarea value={task.task} onChange={(event) => updateTask(task.id, { task: event.target.value })} />
                      </td>
                      <td>
                        <select
                          value={task.backendStatus}
                          className={`status-select ${statusClassName(task.backendStatus)}`}
                          onChange={(event) => updateTask(task.id, { backendStatus: event.target.value })}
                        >
                          {selected.statuses.map((status) => <option key={status.value} value={status.name}>{status.name}</option>)}
                        </select>
                      </td>
                      <td>
                        <input
                          className="number-input"
                          value={task.backendEstimateDays ?? ''}
                          onChange={(event) => updateTask(task.id, { backendEstimateDays: numericValue(event.target.value) })}
                        />
                      </td>
                      <td>
                        <select
                          value={task.frontendStatus}
                          className={`status-select ${statusClassName(task.frontendStatus)}`}
                          onChange={(event) => updateTask(task.id, { frontendStatus: event.target.value })}
                        >
                          {selected.statuses.map((status) => <option key={status.value} value={status.name}>{status.name}</option>)}
                        </select>
                      </td>
                      <td>
                        <input
                          className="number-input"
                          value={task.frontendEstimateDays ?? ''}
                          onChange={(event) => updateTask(task.id, { frontendEstimateDays: numericValue(event.target.value) })}
                        />
                      </td>
                      <td>
                        <div className="computed-status-cell">
                          <span className={`status-badge ${statusClassName(computeOverallStatus(task.backendStatus, task.frontendStatus))}`}>
                            {computeOverallStatus(task.backendStatus, task.frontendStatus)}
                          </span>
                          {missingCounterpartLabel(task.backendStatus, task.frontendStatus) && (
                            <small>{missingCounterpartLabel(task.backendStatus, task.frontendStatus)}</small>
                          )}
                        </div>
                      </td>
                      <td>
                        <textarea
                          value={task.note1 || ''}
                          onChange={(event) => updateTask(task.id, { note1: event.target.value })}
                          placeholder="Nota 1"
                        />
                      </td>
                      <td>
                        <button className="icon-button danger" onClick={() => deleteTask(task.id)} title="Elimina task">
                          <Trash2 size={16} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="table-footer">
              <button className="btn btn-outline" onClick={addTask}><Plus size={16} /> Aggiungi task</button>
            </div>
          </>
        )}
      </section>
    </div>
  );
}
