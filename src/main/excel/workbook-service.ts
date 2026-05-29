import ExcelJS from 'exceljs';
import type {
  ProjectFile,
  ProjectTask,
  ProjectWorkbook,
  StatusOption,
  WorkbookFormat,
} from '../../renderer/types/project';
import { aggregateStatuses, computeOverallStatus, normalizeStatusName } from '../../shared/status';

export const DEFAULT_STATUSES: StatusOption[] = [
  { value: 0, name: 'Non iniziato' },
  { value: 1, name: 'Work in progress' },
  { value: 2, name: 'Da verificare' },
  { value: 3, name: 'Concluso' },
  { value: 4, name: 'In Pausa' },
  { value: 5, name: 'Rimandato' },
  { value: 6, name: 'Da Definire' },
];

type WorkbookLoadBuffer = Parameters<ExcelJS.Workbook['xlsx']['load']>[0];

function toWorkbookLoadBuffer(content: Buffer): WorkbookLoadBuffer {
  const arrayBuffer = content.buffer.slice(content.byteOffset, content.byteOffset + content.byteLength);
  return arrayBuffer as WorkbookLoadBuffer;
}

function cellText(value: ExcelJS.CellValue | undefined): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'object') {
    if ('result' in value && value.result !== undefined) return cellText(value.result as ExcelJS.CellValue);
    if ('text' in value && value.text !== undefined) return String(value.text);
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text).join('');
    }
    if ('hyperlink' in value && 'text' in value) return String(value.text);
  }
  return String(value).trim();
}

function cellNumber(value: ExcelJS.CellValue | undefined): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const text = cellText(value).replace(',', '.').trim();
  if (!text) return null;
  const parsed = Number.parseFloat(text);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeKey(value: string): string {
  return value.toLowerCase().replace(/\s+/g, ' ').trim();
}

export function normalizeStatus(value: ExcelJS.CellValue | string | number | undefined): string {
  const text = normalizeKey(cellText(value as ExcelJS.CellValue));
  const known = normalizeStatusName(text);
  const direct = DEFAULT_STATUSES.find((status) => normalizeKey(status.name) === text);
  return direct?.name || known;
}

function statusValue(statusName: string, statuses: StatusOption[]): number {
  const status = statuses.find((item) => normalizeKey(item.name) === normalizeKey(statusName));
  return status?.value ?? 6;
}

function statusList(workbook: ExcelJS.Workbook): StatusOption[] {
  for (const sheet of workbook.worksheets) {
    const first = normalizeKey(cellText(sheet.getRow(1).getCell(1).value));
    const second = normalizeKey(cellText(sheet.getRow(1).getCell(2).value));
    if ((first === 'valore' && second === 'nome') || first === 'stati') {
      const statuses: StatusOption[] = [];
      for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const row = sheet.getRow(rowNumber);
        const value = cellNumber(row.getCell(1).value);
        const name = cellText(row.getCell(2).value);
        if (value !== null && name) statuses.push({ value, name });
      }
      if (statuses.length > 0) return statuses;
    }
  }
  return DEFAULT_STATUSES;
}

function detectFormat(workbook: ExcelJS.Workbook): { sheet: ExcelJS.Worksheet; format: WorkbookFormat } {
  const piano = workbook.getWorksheet('Piano lavori');
  if (piano) {
    const headers = [1, 2, 3, 4, 5, 6, 7, 8].map((col) =>
      normalizeKey(cellText(piano.getRow(4).getCell(col).value))
    );
    if (headers.includes('task') && headers.includes('stato backend')) {
      return { sheet: piano, format: 'piano-lavori' };
    }
  }

  const first = workbook.worksheets[0];
  if (!first) throw new Error('Workbook Excel vuoto.');
  return { sheet: first, format: 'legacy-status' };
}

function parsePianoSheet(sheet: ExcelJS.Worksheet): ProjectTask[] {
  const tasks: ProjectTask[] = [];

  for (let rowNumber = 5; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const area = cellText(row.getCell(1).value);
    const task = cellText(row.getCell(2).value);
    const hasData = [1, 2, 3, 5, 7, 9, 10].some((col) => cellText(row.getCell(col).value));
    if (!hasData || !task) continue;

    const backendStatus = normalizeStatus(row.getCell(3).value);
    const frontendStatus = normalizeStatus(row.getCell(5).value);
    const taskId = cellText(row.getCell(14).value) || `row-${rowNumber}`;
    const parentId = cellText(row.getCell(15).value) || null;
    const level = cellNumber(row.getCell(16).value) ?? (parentId ? 1 : 0);
    tasks.push({
      id: taskId,
      parentId,
      level,
      rowNumber,
      area,
      task,
      backendStatus,
      backendEstimateDays: cellNumber(row.getCell(4).value),
      frontendStatus,
      frontendEstimateDays: cellNumber(row.getCell(6).value),
      overallStatus: computeOverallStatus(backendStatus, frontendStatus),
      totalEstimateDays: cellNumber(row.getCell(8).value),
      note1: cellText(row.getCell(9).value),
      note2: cellText(row.getCell(10).value),
    });
  }

  return tasks;
}

function parseLegacySheet(sheet: ExcelJS.Worksheet): ProjectTask[] {
  const tasks: ProjectTask[] = [];
  let currentPhase = '';
  let currentArea = '';

  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const phase = cellText(row.getCell(1).value);
    const area = cellText(row.getCell(2).value);
    const task = cellText(row.getCell(3).value);

    if (phase && normalizeKey(phase) !== 'fase') currentPhase = phase;
    if (area && area !== '-' && normalizeKey(area) !== 'sottocategoria') currentArea = area;
    if (!task || normalizeKey(task) === 'task') continue;

    const backendStatus = normalizeStatus(row.getCell(4).value);
    const frontendStatus = normalizeStatus(row.getCell(5).value);
    tasks.push({
      id: `row-${rowNumber}`,
      parentId: null,
      level: 0,
      rowNumber,
      phase: currentPhase,
      area: currentArea || area || '-',
      task,
      backendStatus,
      backendEstimateDays: null,
      frontendStatus,
      frontendEstimateDays: null,
      overallStatus: computeOverallStatus(backendStatus, frontendStatus),
      totalEstimateDays: null,
      note1: cellText(row.getCell(7).value),
      note2: cellText(row.getCell(8).value),
    });
  }

  return tasks;
}

export async function parseProjectWorkbook(content: Buffer, file: ProjectFile): Promise<ProjectWorkbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(toWorkbookLoadBuffer(content));

  const { sheet, format } = detectFormat(workbook);
  const statuses = statusList(workbook);
  const tasks = format === 'piano-lavori' ? parsePianoSheet(sheet) : parseLegacySheet(sheet);

  return {
    file,
    sheetName: sheet.name,
    format,
    statuses,
    tasks,
    loadedAt: new Date().toISOString(),
  };
}

function safeMerge(sheet: ExcelJS.Worksheet, range: string): void {
  try {
    sheet.mergeCells(range);
  } catch {
    // Existing templates may already contain the merge.
  }
}

function styleHeader(sheet: ExcelJS.Worksheet): void {
  sheet.properties.defaultRowHeight = 22;
  sheet.views = [{ state: 'frozen', ySplit: 4 }];

  sheet.columns = [
    { key: 'area', width: 22 },
    { key: 'task', width: 42 },
    { key: 'backendStatus', width: 20 },
    { key: 'backendEstimateDays', width: 16 },
    { key: 'frontendStatus', width: 20 },
    { key: 'frontendEstimateDays', width: 16 },
    { key: 'overallStatus', width: 20 },
    { key: 'totalEstimateDays', width: 14 },
    { key: 'note1', width: 28 },
    { key: 'note2', width: 52 },
    { key: 'backendValue', width: 14, hidden: true },
    { key: 'frontendValue', width: 14, hidden: true },
    { key: 'statusValue', width: 14, hidden: true },
    { key: 'taskId', width: 18, hidden: true },
    { key: 'parentId', width: 18, hidden: true },
    { key: 'level', width: 10, hidden: true },
  ];

  safeMerge(sheet, 'A1:J1');
  safeMerge(sheet, 'A2:J2');
  safeMerge(sheet, 'C3:D3');
  safeMerge(sheet, 'E3:F3');

  sheet.getRow(1).height = 30;
  sheet.getCell('A1').font = { bold: true, size: 18, color: { argb: 'FFFFFFFF' } };
  sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF151515' } };
  sheet.getCell('A1').alignment = { vertical: 'middle' };

  sheet.getCell('A2').font = { italic: true, size: 10, color: { argb: 'FF666666' } };
  sheet.getCell('A2').alignment = { wrapText: true };

  sheet.getRow(3).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(3).alignment = { horizontal: 'center' };
  sheet.getCell('C3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF243B53' } };
  sheet.getCell('E3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3D2C54' } };

  const header = sheet.getRow(4);
  header.height = 24;
  header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  header.alignment = { vertical: 'middle', wrapText: true };
  header.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF222222' } };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF555555' } } };
  });
}

function ensureListsSheet(workbook: ExcelJS.Workbook, statuses: StatusOption[]): ExcelJS.Worksheet {
  let sheet = workbook.getWorksheet('Liste');
  if (!sheet) sheet = workbook.addWorksheet('Liste');
  sheet.spliceRows(1, sheet.rowCount);
  sheet.columns = [{ width: 10 }, { width: 22 }];
  sheet.getRow(1).values = ['Valore', 'Nome'];
  statuses.forEach((status, index) => {
    sheet.getRow(index + 2).values = [status.value, status.name];
  });
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF222222' } };
  });
  return sheet;
}

function writePianoRows(
  sheet: ExcelJS.Worksheet,
  tasks: ProjectTask[],
  statuses: StatusOption[],
  initialRows: number,
  projectName?: string
): void {
  sheet.spliceRows(1, sheet.rowCount);

  sheet.getRow(1).values = [projectName ? `Piano lavori - ${projectName}` : 'Piano lavori'];
  sheet.getRow(2).values = [
    'Gli stati sono selezionabili da elenco. Il nome resta visibile; il valore numerico viene calcolato nelle colonne tecniche K:M.',
  ];
  sheet.getRow(3).values = [null, null, 'Backend', null, 'Frontend'];
  sheet.getRow(4).values = [
    'Ambito',
    'Task',
    'Stato backend',
    'Stima backend (gg)',
    'Stato frontend',
    'Stima frontend (gg)',
    'Stato',
    'Stima totale',
    'Nota 1',
    'Nota 2',
    'Backend valore',
    'Frontend valore',
    'Stato valore',
    'Task ID',
    'Parent ID',
    'Livello',
  ];

  styleHeader(sheet);

  const totalRows = Math.max(initialRows, tasks.length || 1);
  const statusNames = statuses.map((status) => status.name).join(',');
  const rolledTasks = rollupParentTasks(tasks);

  for (let index = 0; index < totalRows; index += 1) {
    const rowNumber = index + 5;
    const task = rolledTasks[index];
    const row = sheet.getRow(rowNumber);
    const backendEstimate = task?.backendEstimateDays ?? null;
    const frontendEstimate = task?.frontendEstimateDays ?? null;
    const backendStatus = task?.backendStatus || 'Da Definire';
    const frontendStatus = task?.frontendStatus || 'Da Definire';
    const overallStatus = task ? computeOverallStatus(backendStatus, frontendStatus) : 'Da Definire';

    row.values = [
      task?.area || null,
      task?.task || null,
      task ? backendStatus : null,
      backendEstimate,
      task ? frontendStatus : null,
      frontendEstimate,
      task ? overallStatus : null,
      {
        formula: `IF(AND(D${rowNumber}="",F${rowNumber}=""),"",N(D${rowNumber})+N(F${rowNumber}))`,
        result: (backendEstimate || 0) + (frontendEstimate || 0) || undefined,
      },
      task?.note1 || null,
      task?.note2 || null,
      task ? statusValue(backendStatus, statuses) : null,
      task ? statusValue(frontendStatus, statuses) : null,
      task ? statusValue(overallStatus, statuses) : null,
      task?.id || null,
      task?.parentId || null,
      task?.level || (task?.parentId ? 1 : 0),
    ];

    [3, 5].forEach((col) => {
      row.getCell(col).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [`"${statusNames}"`],
      };
    });

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.alignment = {
        vertical: 'top',
        wrapText: colNumber === 2 || colNumber === 9 || colNumber === 10,
        indent: colNumber === 2 && task?.parentId ? 1 : 0,
      };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFE6E6E6' } } };
      if ([4, 6, 8].includes(colNumber)) {
        cell.numFmt = '0.0';
      }
    });
  }
}

function rollupParentTasks(tasks: ProjectTask[]): ProjectTask[] {
  const childrenByParent = new Map<string, ProjectTask[]>();
  tasks.forEach((task) => {
    if (!task.parentId) return;
    const children = childrenByParent.get(task.parentId) || [];
    children.push(task);
    childrenByParent.set(task.parentId, children);
  });

  return tasks.map((task) => {
    const children = childrenByParent.get(task.id) || [];
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
  });
}

function legacyToken(status: string): string {
  switch (normalizeStatus(status)) {
    case 'Concluso':
      return 'y';
    case 'Work in progress':
      return 'wip';
    case 'Da verificare':
      return 'vr';
    case 'In Pausa':
      return 'pause';
    case 'Rimandato':
      return 'rimandato';
    case 'Non iniziato':
      return '0';
    default:
      return '-';
  }
}

function writeLegacyRows(sheet: ExcelJS.Worksheet, tasks: ProjectTask[]): void {
  sheet.spliceRows(1, sheet.rowCount);
  sheet.columns = [
    { width: 12 },
    { width: 24 },
    { width: 48 },
    { width: 14 },
    { width: 14 },
    { width: 18 },
    { width: 24 },
    { width: 52 },
  ];
  sheet.getRow(1).values = ['Fase', 'Sottocategoria', 'Task', 'BackEnd', 'FrontEnd', 'Stato', 'Nota 1', 'Nota 2'];
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF222222' } };
  });

  tasks.forEach((task, index) => {
    const row = sheet.getRow(index + 2);
    row.values = [
      task.phase || null,
      task.area || null,
      task.task,
      legacyToken(task.backendStatus),
      legacyToken(task.frontendStatus),
      computeOverallStatus(task.backendStatus, task.frontendStatus),
      task.note1 || null,
      task.note2 || null,
    ];
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.alignment = { vertical: 'top', wrapText: true };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFE6E6E6' } } };
    });
  });
}

export async function writeProjectWorkbook(
  content: Buffer,
  project: Pick<ProjectWorkbook, 'format' | 'sheetName' | 'tasks' | 'statuses'>,
  initialRows: number,
  projectName?: string
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(toWorkbookLoadBuffer(content));

  if (project.format === 'legacy-status') {
    const sheet = workbook.getWorksheet(project.sheetName) || workbook.worksheets[0] || workbook.addWorksheet('Foglio1');
    writeLegacyRows(sheet, project.tasks);
  } else {
    const sheet = workbook.getWorksheet(project.sheetName) || workbook.getWorksheet('Piano lavori') || workbook.addWorksheet('Piano lavori');
    ensureListsSheet(workbook, project.statuses.length ? project.statuses : DEFAULT_STATUSES);
    writePianoRows(sheet, project.tasks, project.statuses.length ? project.statuses : DEFAULT_STATUSES, initialRows, projectName);
  }

  const output = await workbook.xlsx.writeBuffer();
  return Buffer.from(output);
}

export async function createProjectWorkbook(
  projectName: string,
  tasks: ProjectTask[],
  initialRows: number,
  templateContent?: Buffer
): Promise<Buffer> {
  if (templateContent) {
    const parsed = await parseProjectWorkbook(templateContent, { id: 'template', name: 'template.xlsx' });
    return writeProjectWorkbook(
      templateContent,
      {
        ...parsed,
        tasks,
        statuses: parsed.statuses.length ? parsed.statuses : DEFAULT_STATUSES,
        format: parsed.format,
      },
      initialRows,
      projectName
    );
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Project Step Manager';
  workbook.created = new Date();
  const sheet = workbook.addWorksheet('Piano lavori');
  ensureListsSheet(workbook, DEFAULT_STATUSES);
  writePianoRows(sheet, tasks, DEFAULT_STATUSES, initialRows, projectName);
  const output = await workbook.xlsx.writeBuffer();
  return Buffer.from(output);
}
