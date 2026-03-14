import ExcelJS from 'exceljs';

// Extend FileSystemDirectoryHandle with Chrome-specific permission methods
declare global {
  interface FileSystemHandle {
    queryPermission(desc?: { mode?: string }): Promise<string>;
    requestPermission(desc?: { mode?: string }): Promise<string>;
  }
}

// ── IndexedDB helpers for persisting the directory handle ──

const DB_NAME = 'L10MeetingTool';
const STORE_NAME = 'config';
const HANDLE_KEY = 'rootDirHandle';

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE_NAME);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function getStoredHandle(): Promise<FileSystemDirectoryHandle | null> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).get(HANDLE_KEY);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

async function storeHandle(handle: FileSystemDirectoryHandle): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put(handle, HANDLE_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

// ── Directory handle state ──

let _rootHandle: FileSystemDirectoryHandle | null = null;

/** Check if we have a stored directory handle (may or may not have permission) */
export async function hasStoredFolder(): Promise<'granted' | 'prompt' | false> {
  const handle = await getStoredHandle();
  if (!handle) return false;
  const perm = await handle.queryPermission({ mode: 'readwrite' });
  return perm === 'granted' ? 'granted' : 'prompt';
}

/** Try to restore a previously selected folder (may need user gesture for permission) */
export async function restoreFolder(): Promise<boolean> {
  const handle = await getStoredHandle();
  if (!handle) return false;
  const perm = await handle.requestPermission({ mode: 'readwrite' });
  if (perm !== 'granted') return false;
  _rootHandle = handle;
  await ensureStructure();
  return true;
}

/** Prompt user to pick a folder */
export async function pickFolder(): Promise<boolean> {
  try {
    const handle = await (window as any).showDirectoryPicker({ mode: 'readwrite' });
    _rootHandle = handle;
    await storeHandle(handle);
    await ensureStructure();
    return true;
  } catch {
    return false;
  }
}

/** Forget the stored folder */
export async function forgetFolder(): Promise<void> {
  _rootHandle = null;
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).delete(HANDLE_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export function getFolderName(): string {
  return _rootHandle?.name ?? '';
}

/** Load logo.png from the data folder as an object URL, or null if not found */
export async function loadLogo(): Promise<string | null> {
  if (!_rootHandle) return null;
  try {
    const fileHandle = await _rootHandle.getFileHandle('logo.png');
    const file = await fileHandle.getFile();
    return URL.createObjectURL(file);
  } catch {
    return null;
  }
}

/** Save an image file as logo.png in the data folder */
export async function saveLogo(file: File): Promise<string> {
  if (!_rootHandle) throw new Error('No folder selected');
  const fileHandle = await _rootHandle.getFileHandle('logo.png', { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(file);
  await writable.close();
  // Return object URL for immediate use
  const saved = await fileHandle.getFile();
  return URL.createObjectURL(saved);
}

/** Delete logo.png from the data folder */
export async function deleteLogo(): Promise<void> {
  if (!_rootHandle) return;
  try {
    await _rootHandle.removeEntry('logo.png');
  } catch { /* doesn't exist */ }
}

async function ensureStructure(): Promise<void> {
  if (!_rootHandle) return;
  await _rootHandle.getDirectoryHandle('Departments', { create: true });
}

// ── Internal helpers ──

async function getDeptHandle(name: string, create = false): Promise<FileSystemDirectoryHandle> {
  const deps = await _rootHandle!.getDirectoryHandle('Departments');
  return deps.getDirectoryHandle(name, { create });
}

async function getMeetingsHandle(deptName: string, create = false): Promise<FileSystemDirectoryHandle> {
  const dept = await getDeptHandle(deptName, create);
  return dept.getDirectoryHandle('meetings', { create });
}

async function readTextFile(dirHandle: FileSystemDirectoryHandle, name: string): Promise<string> {
  try {
    const fileHandle = await dirHandle.getFileHandle(name);
    const file = await fileHandle.getFile();
    return await file.text();
  } catch {
    return '';
  }
}

async function writeTextFile(dirHandle: FileSystemDirectoryHandle, name: string, content: string): Promise<void> {
  const fileHandle = await dirHandle.getFileHandle(name, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(content);
  await writable.close();
}

async function readExcelData(dirHandle: FileSystemDirectoryHandle, fileName: string): Promise<Record<string, any> | null> {
  try {
    const fileHandle = await dirHandle.getFileHandle(fileName);
    const file = await fileHandle.getFile();
    const buffer = await file.arrayBuffer();

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    // Try _data sheet first (lossless)
    const dataWs = wb.getWorksheet('_data');
    if (dataWs) {
      const raw = dataWs.getCell('A1').value;
      if (raw && typeof raw === 'string') {
        try { return JSON.parse(raw); } catch { /* fall through */ }
      }
    }

    // Fallback: read from L10 Meeting sheet
    const ws = wb.getWorksheet('L10 Meeting');
    if (ws) return readWorkbookToJson(ws);
    return null;
  } catch {
    return null;
  }
}

async function writeExcelData(dirHandle: FileSystemDirectoryHandle, fileName: string, data: Record<string, any>): Promise<void> {
  const wb = new ExcelJS.Workbook();

  // Try to read existing file as base
  try {
    const fileHandle = await dirHandle.getFileHandle(fileName);
    const file = await fileHandle.getFile();
    const buffer = await file.arrayBuffer();
    await wb.xlsx.load(buffer);
  } catch {
    wb.addWorksheet('L10 Meeting');
  }

  const ws = wb.getWorksheet('L10 Meeting');
  if (ws) writeJsonToWorkbook(ws, data);

  let dataWs = wb.getWorksheet('_data');
  if (!dataWs) dataWs = wb.addWorksheet('_data');
  dataWs.getCell('A1').value = JSON.stringify(data);

  const buffer = await wb.xlsx.writeBuffer();
  const fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(buffer);
  await writable.close();
}

// ── Public API (replaces all fetch calls) ──

export async function getDepartments(): Promise<{ name: string; peopleCount: number }[]> {
  if (!_rootHandle) return [];
  try {
    const deps = await _rootHandle.getDirectoryHandle('Departments');
    const results: { name: string; peopleCount: number }[] = [];
    for await (const entry of (deps as any).values()) {
      if (entry.kind !== 'directory') continue;
      const content = await readTextFile(entry, 'people.txt');
      const peopleCount = content.trim() ? content.trim().split('\n').filter(Boolean).length : 0;
      results.push({ name: entry.name, peopleCount });
    }
    return results;
  } catch {
    return [];
  }
}

export async function createDepartment(name: string): Promise<{ ok: boolean; error?: string }> {
  if (!_rootHandle) return { ok: false, error: 'No folder selected' };
  try {
    const deps = await _rootHandle.getDirectoryHandle('Departments');
    // Check if exists
    try {
      await deps.getDirectoryHandle(name);
      return { ok: false, error: 'Already exists' };
    } catch { /* doesn't exist, good */ }
    const dept = await deps.getDirectoryHandle(name, { create: true });
    await dept.getDirectoryHandle('meetings', { create: true });
    await writeTextFile(dept, 'people.txt', '');
    return { ok: true };
  } catch (err: any) {
    return { ok: false, error: err.message };
  }
}

export async function renameDepartment(oldName: string, newName: string): Promise<{ ok: boolean; error?: string }> {
  if (!_rootHandle) return { ok: false, error: 'No folder selected' };
  // File System Access API doesn't support rename directly.
  // We need to copy everything to a new directory and delete the old one.
  try {
    const deps = await _rootHandle.getDirectoryHandle('Departments');
    // Check new name doesn't exist
    try {
      await deps.getDirectoryHandle(newName);
      return { ok: false, error: 'Target name already exists' };
    } catch { /* good */ }

    const oldDept = await deps.getDirectoryHandle(oldName);
    const newDept = await deps.getDirectoryHandle(newName, { create: true });

    // Copy people.txt
    const peopleContent = await readTextFile(oldDept, 'people.txt');
    await writeTextFile(newDept, 'people.txt', peopleContent);

    // Copy meetings folder
    try {
      const oldMeetings = await oldDept.getDirectoryHandle('meetings');
      const newMeetings = await newDept.getDirectoryHandle('meetings', { create: true });
      for await (const entry of (oldMeetings as any).values()) {
        if (entry.kind !== 'file') continue;
        const file = await entry.getFile();
        const buffer = await file.arrayBuffer();
        const newFile = await newMeetings.getFileHandle(entry.name, { create: true });
        const writable = await newFile.createWritable();
        await writable.write(buffer);
        await writable.close();
      }
    } catch { /* no meetings dir */ }

    // Delete old department
    await deps.removeEntry(oldName, { recursive: true });
    return { ok: true };
  } catch (err: any) {
    return { ok: false, error: err.message };
  }
}

export async function deleteDepartment(name: string): Promise<{ ok: boolean }> {
  if (!_rootHandle) return { ok: false };
  try {
    const deps = await _rootHandle.getDirectoryHandle('Departments');
    await deps.removeEntry(name, { recursive: true });
    return { ok: true };
  } catch {
    return { ok: false };
  }
}

export async function getPeople(deptName: string): Promise<string[]> {
  if (!_rootHandle) return [];
  try {
    const dept = await getDeptHandle(deptName);
    const content = await readTextFile(dept, 'people.txt');
    return content.trim() ? content.trim().split('\n').map(s => s.trim()).filter(Boolean) : [];
  } catch {
    return [];
  }
}

export async function savePeople(deptName: string, people: string[]): Promise<void> {
  if (!_rootHandle) return;
  try {
    const dept = await getDeptHandle(deptName, true);
    await writeTextFile(dept, 'people.txt', people.join('\n'));
  } catch { /* silent */ }
}

export async function getMeetings(deptName: string): Promise<{ id: string; date: string; lastSaved: string; avgRating: number }[]> {
  if (!_rootHandle) return [];
  try {
    const meetings = await getMeetingsHandle(deptName);
    const results: { id: string; date: string; lastSaved: string; avgRating: number }[] = [];

    for await (const entry of (meetings as any).values()) {
      if (entry.kind !== 'file' || !entry.name.endsWith('.xlsx') || entry.name.startsWith('~$')) continue;
      const id = entry.name.replace('.xlsx', '');
      const dateMatch = entry.name.match(/(\d{4}-\d{2}-\d{2}(-\d+)?)/);
      const date = dateMatch ? dateMatch[1] : id;

      let avgRating = 0;
      let lastSaved = '';
      try {
        const file = await entry.getFile();
        lastSaved = new Date(file.lastModified).toISOString();
        const buffer = await file.arrayBuffer();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const dataWs = wb.getWorksheet('_data');
        if (dataWs) {
          const raw = dataWs.getCell('A1').value;
          if (raw && typeof raw === 'string') {
            const data = JSON.parse(raw);
            if (data.lastSaved) lastSaved = data.lastSaved;
            const vals = (data.ratingValues || []) as string[];
            let sum = 0, count = 0;
            vals.forEach((v: string) => { const n = parseInt(v); if (n > 0) { sum += n; count++; } });
            if (count > 0) avgRating = sum / count;
          }
        }
      } catch { /* skip */ }

      results.push({ id, date, lastSaved, avgRating });
    }

    results.sort((a, b) => b.date.localeCompare(a.date));
    return results;
  } catch {
    return [];
  }
}

export async function getMeetingData(deptName: string, meetingId: string): Promise<Record<string, any> | null> {
  if (!_rootHandle) return null;
  try {
    const meetings = await getMeetingsHandle(deptName);
    return await readExcelData(meetings, `${meetingId}.xlsx`);
  } catch {
    return null;
  }
}

export async function createMeeting(deptName: string, data: Record<string, any>): Promise<{ id: string } | null> {
  if (!_rootHandle) return null;
  try {
    const meetings = await getMeetingsHandle(deptName, true);
    const today = new Date().toISOString().split('T')[0];
    let baseName = `L10_${deptName}_${today}`;
    let fileName = `${baseName}.xlsx`;
    let suffix = 1;

    // Check for existing files
    const existingNames = new Set<string>();
    for await (const entry of (meetings as any).values()) {
      if (entry.kind === 'file') existingNames.add(entry.name);
    }
    while (existingNames.has(fileName)) {
      fileName = `${baseName}-${suffix}.xlsx`;
      suffix++;
    }

    const id = fileName.replace('.xlsx', '');
    data.createdAt = new Date().toISOString();
    data.lastSaved = new Date().toISOString();

    await writeExcelData(meetings, fileName, data);
    return { id };
  } catch {
    return null;
  }
}

export async function saveMeeting(deptName: string, meetingId: string, data: Record<string, any>): Promise<boolean> {
  if (!_rootHandle) return false;
  try {
    const meetings = await getMeetingsHandle(deptName, true);
    data.lastSaved = new Date().toISOString();
    await writeExcelData(meetings, `${meetingId}.xlsx`, data);
    return true;
  } catch {
    return false;
  }
}

export async function deleteMeeting(deptName: string, meetingId: string): Promise<boolean> {
  if (!_rootHandle) return false;
  try {
    const meetings = await getMeetingsHandle(deptName);
    await meetings.removeEntry(`${meetingId}.xlsx`);
    return true;
  } catch {
    return false;
  }
}

export async function importMeetingFile(deptName: string, fileData: ArrayBuffer): Promise<{ id: string } | null> {
  if (!_rootHandle) return null;
  try {
    const meetings = await getMeetingsHandle(deptName, true);
    const today = new Date().toISOString().split('T')[0];
    let baseName = `L10_${deptName}_${today}`;
    let fileName = `${baseName}.xlsx`;
    let suffix = 1;

    const existingNames = new Set<string>();
    for await (const entry of (meetings as any).values()) {
      if (entry.kind === 'file') existingNames.add(entry.name);
    }
    while (existingNames.has(fileName)) {
      fileName = `${baseName}-${suffix}.xlsx`;
      suffix++;
    }

    const id = fileName.replace('.xlsx', '');
    const fileHandle = await meetings.getFileHandle(fileName, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(fileData);
    await writable.close();
    return { id };
  } catch {
    return null;
  }
}

export async function downloadMeetingExcel(deptName: string, meetingId: string): Promise<void> {
  if (!_rootHandle) return;
  try {
    const meetings = await getMeetingsHandle(deptName);
    const fileHandle = await meetings.getFileHandle(`${meetingId}.xlsx`);
    const file = await fileHandle.getFile();
    const url = URL.createObjectURL(file);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${meetingId}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch { /* silent */ }
}

// ── Excel helpers (ported from server/index.ts) ──

function stripEmoji(s: string): string {
  return s.replace(/[\u{1F000}-\u{1FFFF}]|[\u{2600}-\u{27BF}]|[\u{FE00}-\u{FE0F}]|[\u{1F900}-\u{1F9FF}]|[\u{200D}]|[\u{20E3}]|[\u{E0020}-\u{E007F}]/gu, '').trim();
}

function cellStr(ws: ExcelJS.Worksheet, ref: string): string {
  const v = ws.getCell(ref).value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && 'result' in (v as any)) {
    const r = (v as any).result;
    if (r === null || r === undefined) return '';
    if (r instanceof Date) return formatDateCell(r);
    return stripEmoji(String(r));
  }
  if (v instanceof Date) return formatDateCell(v);
  return stripEmoji(String(v));
}

function formatDateCell(d: Date): string {
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function readWorkbookToJson(ws: ExcelJS.Worksheet): Record<string, any> {
  const c = (ref: string) => cellStr(ws, ref);
  const data: Record<string, any> = {
    meta: {
      team: c('B2'), date: c('E2'), facilitator: c('B3'),
      scribe: c('E3'), start: c('B4'), end: c('E4'),
    },
    segue: { personal: c('B8'), professional: c('B9') },
  };

  data.scorecardTable = readTable(ws, 14, 7, ['A', 'B', 'C', 'D', 'E', 'F']);
  data.okrReviewTable = readTable(ws, 26, 6, ['A', 'B', 'C', 'D', 'E', 'F']);
  data.headlinesTable = readTable(ws, 37, 6, ['A', 'B', 'C', 'D', 'E', 'F']);
  data.todoReviewTable = readTable(ws, 47, 7, ['A', 'B', 'C', 'D', 'E', 'F']);
  data.issuesListTable = readTable(ws, 60, 16, ['A', 'B', 'C', 'D', 'E', 'F']);

  const idsOffset = c('A79').startsWith('Issue:') ? 1 : 0;
  const idsBlocks: any[] = [];
  for (let bi = 0; bi < 10; bi++) {
    const base = 77 + idsOffset + bi * 9;
    const fields = [c(`B${base + 1}`), c(`B${base + 2}`), c(`B${base + 3}`)];
    const isPlaceholder = (s: string) => !s || s.startsWith('Describe the real') || s.startsWith("Ask 'why?'") || s.startsWith('Agreed solution');
    if (isPlaceholder(fields[0]) && isPlaceholder(fields[1]) && isPlaceholder(fields[2])) continue;
    const todoStart = idsOffset ? base + 5 : base + 4;
    idsBlocks.push({ fields, todos: readTable(ws, todoStart, 5, ['A', 'B', 'C', 'D', 'E', 'F']) });
  }
  data.idsBlocks = idsBlocks;

  const newTodoStart = 171 + idsOffset;
  const cascadingStart = 184 + idsOffset;
  const ratingStart = 192 + idsOffset;

  data.newTodoTable = readTable(ws, newTodoStart, 11, ['A', 'B', 'C', 'D', 'E', 'F']);
  data.cascadingTable = readTable(ws, cascadingStart, 6, ['A', 'B', 'C', 'D', 'E', 'F']);

  const ratingTable: string[][] = [];
  const ratingValues: string[] = [];
  for (let i = 0; i < 10; i++) {
    const r = ratingStart + i;
    const name = c(`A${r}`);
    const rating = c(`B${r}`);
    const comment = c(`C${r}`);
    if (!name || name.startsWith('Average')) break;
    ratingTable.push([name, '', comment]);
    ratingValues.push(rating || '0');
  }
  data.ratingTable = ratingTable;
  data.ratingValues = ratingValues;

  return data;
}

function readTable(ws: ExcelJS.Worksheet, startRow: number, maxRows: number, cols: string[]): string[][] {
  const rows: string[][] = [];
  for (let i = 0; i < maxRows; i++) {
    const r = startRow + i;
    const row = cols.map(col => cellStr(ws, `${col}${r}`));
    if (row.every(v => !v)) continue;
    rows.push(row);
  }
  return rows;
}

function writeJsonToWorkbook(ws: ExcelJS.Worksheet, data: Record<string, any>): void {
  const meta = data.meta || {};
  ws.getCell('B2').value = meta.team || '';
  ws.getCell('E2').value = meta.date || '';
  ws.getCell('B3').value = meta.facilitator || '';
  ws.getCell('E3').value = meta.scribe || '';
  ws.getCell('B4').value = meta.start || '';
  ws.getCell('E4').value = meta.end || '';

  const segue = data.segue || {};
  ws.getCell('B8').value = segue.personal || '';
  ws.getCell('B9').value = segue.professional || '';

  writeTable(ws, data.scorecardTable, 14, 7, ['A', 'B', 'C', 'D', 'E', 'F']);
  writeTable(ws, data.okrReviewTable, 26, 6, ['A', 'B', 'C', 'D', 'E', 'F']);
  writeTable(ws, data.headlinesTable, 37, 6, ['A', 'B', 'C', 'D', 'E', 'F']);
  writeTable(ws, data.todoReviewTable, 47, 7, ['A', 'B', 'C', 'D', 'E', 'F']);

  const todos = data.todoReviewTable || [];
  let done = 0;
  todos.forEach((r: string[]) => { if (r[3] === 'Done') done++; });
  ws.getCell('E54').value = `${done} / ${todos.length} done`;

  writeTable(ws, data.issuesListTable, 60, 16, ['A', 'B', 'C', 'D', 'E', 'F']);

  const issueStarts = [77, 86, 95, 104, 113, 122, 131, 140, 149, 158];
  const idsBlocks = data.idsBlocks || [];
  idsBlocks.forEach((block: any, bi: number) => {
    if (bi >= issueStarts.length) return;
    const base = issueStarts[bi];
    const fields = block.fields || [];
    if (fields[0]) ws.getCell(`B${base + 1}`).value = fields[0];
    if (fields[1]) ws.getCell(`B${base + 2}`).value = fields[1];
    if (fields[2]) ws.getCell(`B${base + 3}`).value = fields[2];
    writeTable(ws, block.todos, base + 4, 5, ['A', 'B', 'C', 'D', 'E', 'F']);
  });

  writeTable(ws, data.newTodoTable, 171, 11, ['A', 'B', 'C', 'D', 'E', 'F']);
  writeTable(ws, data.cascadingTable, 184, 6, ['A', 'B', 'C', 'D', 'E', 'F']);

  const ratingTable = data.ratingTable || [];
  const ratingValues = data.ratingValues || [];
  let ratingSum = 0, ratingCount = 0;
  ratingTable.forEach((row: string[], i: number) => {
    if (i > 5) return;
    const r = 192 + i;
    ws.getCell(`A${r}`).value = row[0] || '';
    const rv = parseInt(ratingValues[i]) || 0;
    ws.getCell(`B${r}`).value = rv > 0 ? rv : '';
    ws.getCell(`C${r}`).value = row[2] || '';
    if (rv > 0) { ratingSum += rv; ratingCount++; }
  });
  ws.getCell('B198').value = ratingCount > 0 ? (ratingSum / ratingCount).toFixed(1) : '';

  const validations: { col: string; startRow: number; count: number; options: string[] }[] = [
    { col: 'E', startRow: 14, count: 7, options: ['On Track', 'Off Track', 'At Risk'] },
    { col: 'D', startRow: 26, count: 6, options: ['On Track', 'Off Track', 'At Risk'] },
    { col: 'B', startRow: 37, count: 6, options: ['Customer', 'Employee'] },
    { col: 'D', startRow: 37, count: 6, options: ['Yes', 'No'] },
    { col: 'E', startRow: 37, count: 6, options: ['Yes', 'No'] },
    { col: 'D', startRow: 47, count: 7, options: ['Open', 'Done', 'Carry Over'] },
    { col: 'E', startRow: 47, count: 7, options: ['Yes', 'No'] },
    { col: 'C', startRow: 60, count: 16, options: ['High', 'Medium', 'Low'] },
    { col: 'D', startRow: 60, count: 16, options: ['Open', 'Solved', 'Next Meeting', 'Dropped'] },
    { col: 'F', startRow: 60, count: 16, options: ['Yes', 'No'] },
    { col: 'D', startRow: 171, count: 11, options: ['High', 'Medium', 'Low'] },
    { col: 'E', startRow: 171, count: 11, options: ['Not Started', 'In Progress', 'Done'] },
    { col: 'F', startRow: 184, count: 6, options: ['Yes', 'No'] },
  ];
  for (const v of validations) {
    for (let i = 0; i < v.count; i++) {
      ws.getCell(`${v.col}${v.startRow + i}`).dataValidation = {
        type: 'list', allowBlank: true,
        formulae: [`"${v.options.join(',')}"`],
      };
    }
  }

  for (let bi = 0; bi < issueStarts.length; bi++) {
    const base = issueStarts[bi];
    for (let i = 0; i < 5; i++) {
      ws.getCell(`D${base + 4 + i}`).dataValidation = {
        type: 'list', allowBlank: true, formulae: ['"High,Medium,Low"'],
      };
      ws.getCell(`E${base + 4 + i}`).dataValidation = {
        type: 'list', allowBlank: true, formulae: ['"Not Started,In Progress,Done"'],
      };
    }
  }
}

function writeTable(ws: ExcelJS.Worksheet, rows: string[][] | undefined, startRow: number, maxRows: number, cols: string[]): void {
  if (!rows) return;
  rows.forEach((row, i) => {
    if (i >= maxRows) return;
    const r = startRow + i;
    cols.forEach((col, ci) => {
      ws.getCell(`${col}${r}`).value = row[ci] || '';
    });
  });
}
