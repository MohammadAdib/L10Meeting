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
    if (ws) return readWorkbookToJson(wb, ws);
    return null;
  } catch {
    return null;
  }
}

async function writeExcelData(dirHandle: FileSystemDirectoryHandle, fileName: string, data: Record<string, any>): Promise<void> {
  const wb = new ExcelJS.Workbook();

  // Build all sheets from scratch — dynamic row counts, clean formatting
  buildL10Sheet(wb, data);
  buildScorecardSheet(wb, data);
  buildOkrsSheet(wb, data);

  // Lossless JSON storage (preferred on read)
  const dataWs = wb.addWorksheet('_data');
  dataWs.getCell('A1').value = JSON.stringify(data);

  const buffer = await wb.xlsx.writeBuffer();
  const fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(buffer);
  await writable.close();
}

// ── Public API (replaces all fetch calls) ──

export async function getDepartments(): Promise<{ name: string; peopleCount: number }[]> {
  const ck = cacheKey('departments');
  if (!_rootHandle) return cacheGet(ck) ?? [];
  try {
    const deps = await _rootHandle.getDirectoryHandle('Departments');
    const results: { name: string; peopleCount: number }[] = [];
    for await (const entry of (deps as any).values()) {
      if (entry.kind !== 'directory') continue;
      const content = await readTextFile(entry, 'people.txt');
      const peopleCount = content.trim() ? content.trim().split('\n').filter(Boolean).length : 0;
      results.push({ name: entry.name, peopleCount });
    }
    cacheSet(ck, results);
    return results;
  } catch {
    return cacheGet(ck) ?? [];
  }
}

// ── Cache layer ──
// Stores results keyed by function+args. Write operations invalidate relevant keys.

const _cache = new Map<string, { data: unknown; time: number }>();

function cacheKey(...parts: string[]): string { return parts.join('::'); }

function cacheGet<T>(key: string): T | undefined {
  const entry = _cache.get(key);
  return entry ? entry.data as T : undefined;
}

function cacheSet(key: string, data: unknown): void {
  _cache.set(key, { data, time: Date.now() });
}

function cacheInvalidate(prefix: string): void {
  for (const key of _cache.keys()) {
    if (key.startsWith(prefix)) _cache.delete(key);
  }
}

export function invalidateCache(): void {
  _cache.clear();
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
    cacheInvalidate('departments');
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
    cacheInvalidate('departments');
    cacheInvalidate(cacheKey('people', oldName));
    cacheInvalidate(cacheKey('meetings', oldName));
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
    cacheInvalidate('departments');
    cacheInvalidate(cacheKey('people', name));
    cacheInvalidate(cacheKey('meetings', name));
    return { ok: true };
  } catch {
    return { ok: false };
  }
}

export async function getPeople(deptName: string): Promise<string[]> {
  const ck = cacheKey('people', deptName);
  if (!_rootHandle) return cacheGet(ck) ?? [];
  try {
    const dept = await getDeptHandle(deptName);
    const content = await readTextFile(dept, 'people.txt');
    const result = content.trim() ? content.trim().split('\n').map(s => s.trim()).filter(Boolean) : [];
    cacheSet(ck, result);
    return result;
  } catch {
    return [];
  }
}

export async function savePeople(deptName: string, people: string[]): Promise<void> {
  if (!_rootHandle) return;
  try {
    const dept = await getDeptHandle(deptName, true);
    await writeTextFile(dept, 'people.txt', people.join('\n'));
    cacheSet(cacheKey('people', deptName), [...people]);
    cacheInvalidate('departments'); // people count changed
  } catch { /* silent */ }
}

export async function getMeetings(deptName: string): Promise<{ id: string; date: string; lastSaved: string; avgRating: number }[]> {
  const ck = cacheKey('meetings', deptName);
  const cached = cacheGet<{ id: string; date: string; lastSaved: string; avgRating: number }[]>(ck);
  if (cached) return cached;
  if (!_rootHandle) return [];
  try {
    const meetings = await getMeetingsHandle(deptName);
    const results: { id: string; date: string; lastSaved: string; avgRating: number }[] = [];

    for await (const entry of (meetings as any).values()) {
      if (entry.kind !== 'file' || !entry.name.endsWith('.xlsx') || entry.name.startsWith('~$')) continue;
      const id = entry.name.replace('.xlsx', '');
      const dateMatch = entry.name.match(/(\d{4}-\d{2}-\d{2}(-\d+)?)/);
      const date = dateMatch ? dateMatch[1] : id;

      let lastSaved = '';
      try {
        const file = await entry.getFile();
        lastSaved = new Date(file.lastModified).toISOString();
      } catch { /* skip */ }

      results.push({ id, date, lastSaved, avgRating: 0 });
    }

    results.sort((a, b) => b.date.localeCompare(a.date));
    cacheSet(ck, results);
    return results;
  } catch {
    return [];
  }
}

/** Load ratings for meetings in background (expensive — parses each Excel file) */
export async function loadMeetingRatings(deptName: string): Promise<{ id: string; date: string; lastSaved: string; avgRating: number }[]> {
  const ck = cacheKey('meetings', deptName);
  if (!_rootHandle) return cacheGet(ck) ?? [];
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
    cacheSet(ck, results);
    return results;
  } catch {
    return cacheGet(ck) ?? [];
  }
}

export async function getMeetingData(deptName: string, meetingId: string): Promise<Record<string, any> | null> {
  const ck = cacheKey('meetingData', deptName, meetingId);
  if (!_rootHandle) return cacheGet(ck) ?? null;
  try {
    const meetings = await getMeetingsHandle(deptName);
    const result = await readExcelData(meetings, `${meetingId}.xlsx`);
    cacheSet(ck, result);
    return result;
  } catch {
    return cacheGet(ck) ?? null;
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
    cacheInvalidate(cacheKey('meetings', deptName));
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
    cacheInvalidate(cacheKey('meetingData', deptName, meetingId));
    cacheInvalidate(cacheKey('meetings', deptName));
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
    cacheInvalidate(cacheKey('meetings', deptName));
    cacheInvalidate(cacheKey('meetingData', deptName, meetingId));
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
    cacheInvalidate(cacheKey('meetings', deptName));
    return { id };
  } catch {
    return null;
  }
}


// ── Excel helpers ──

function stripEmoji(s: string): string {
  return s.replace(/[\u{1F000}-\u{1FFFF}]|[\u{2300}-\u{23FF}]|[\u{2600}-\u{27BF}]|[\u{FE00}-\u{FE0F}]|[\u{1F900}-\u{1F9FF}]|[\u{200D}]|[\u{20E3}]|[\u{E0020}-\u{E007F}]|[\u{2700}-\u{27BF}]|[\u{2B50}]|[\u{2705}]|[\u{274C}]/gu, '').trim();
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

// ── Styling constants ──

const BORDER_THIN = { style: 'thin' as const, color: { argb: 'FFBDC3C7' } };
const BORDERS = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN };
const SECTION_BG = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF2C3E50' } };
const SECTION_FONT = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
const HEADER_BG = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFECF0F1' } };
const HEADER_FONT = { bold: true, size: 10 };
const LABEL_FONT = { bold: true, size: 10 };
const SUB_FONT = { bold: true, size: 11, color: { argb: 'FF2C3E50' } };

// ── Write helpers ──

function writeSectionRow(ws: ExcelJS.Worksheet, r: number, title: string, lastCol = 'F'): number {
  ws.mergeCells(`A${r}:${lastCol}${r}`);
  const cell = ws.getCell(`A${r}`);
  cell.value = title;
  cell.font = SECTION_FONT;
  cell.fill = SECTION_BG;
  cell.alignment = { vertical: 'middle' };
  return r + 1;
}

function writeSubHeader(ws: ExcelJS.Worksheet, r: number, title: string, lastCol = 'F'): number {
  ws.mergeCells(`A${r}:${lastCol}${r}`);
  const cell = ws.getCell(`A${r}`);
  cell.value = title;
  cell.font = SUB_FONT;
  return r + 1;
}

function writeHeaders(ws: ExcelJS.Worksheet, r: number, headers: string[], cols: string[]): number {
  cols.forEach((col, i) => {
    if (i >= headers.length) return;
    const cell = ws.getCell(`${col}${r}`);
    cell.value = headers[i];
    cell.font = HEADER_FONT;
    cell.fill = HEADER_BG;
    cell.border = BORDERS;
  });
  return r + 1;
}

function writeRows(ws: ExcelJS.Worksheet, r: number, rows: string[][], cols: string[], validations?: Record<string, string[]>): number {
  if (!rows) return r;
  rows.forEach(row => {
    cols.forEach((col, ci) => {
      const cell = ws.getCell(`${col}${r}`);
      cell.value = row[ci] || '';
      cell.border = BORDERS;
      if (validations?.[col]) {
        cell.dataValidation = { type: 'list', allowBlank: true, formulae: [`"${validations[col].join(',')}"`] };
      }
    });
    r++;
  });
  return r;
}

// ── Build sheets from data (dynamic row counts) ──

function buildL10Sheet(wb: ExcelJS.Workbook, data: Record<string, any>): void {
  const ws = wb.addWorksheet('L10 Meeting');
  const C = ['A', 'B', 'C', 'D', 'E', 'F'];
  [35, 20, 15, 15, 15, 25].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  let r = 1;

  // Title
  ws.mergeCells(`A${r}:F${r}`);
  ws.getCell(`A${r}`).value = 'L10 Meeting';
  ws.getCell(`A${r}`).font = { bold: true, size: 16 };
  r += 2;

  // Meta
  const meta = data.meta || {};
  ws.getCell(`A${r}`).value = 'Team:'; ws.getCell(`A${r}`).font = LABEL_FONT;
  ws.getCell(`B${r}`).value = meta.team || '';
  ws.getCell(`D${r}`).value = 'Date:'; ws.getCell(`D${r}`).font = LABEL_FONT;
  ws.getCell(`E${r}`).value = meta.date || '';
  r++;
  ws.getCell(`A${r}`).value = 'Facilitator:'; ws.getCell(`A${r}`).font = LABEL_FONT;
  ws.getCell(`B${r}`).value = meta.facilitator || '';
  ws.getCell(`D${r}`).value = 'Scribe:'; ws.getCell(`D${r}`).font = LABEL_FONT;
  ws.getCell(`E${r}`).value = meta.scribe || '';
  r++;
  ws.getCell(`A${r}`).value = 'Start Time:'; ws.getCell(`A${r}`).font = LABEL_FONT;
  ws.getCell(`B${r}`).value = meta.start || '';
  ws.getCell(`D${r}`).value = 'End Time:'; ws.getCell(`D${r}`).font = LABEL_FONT;
  ws.getCell(`E${r}`).value = meta.end || '';
  r += 2;

  // 1. SEGUE
  r = writeSectionRow(ws, r, '1. SEGUE');
  ws.getCell(`A${r}`).value = 'Personal Good News:'; ws.getCell(`A${r}`).font = LABEL_FONT;
  ws.getCell(`B${r}`).value = (data.segue || {}).personal || '';
  r++;
  ws.getCell(`A${r}`).value = 'Professional Good News:'; ws.getCell(`A${r}`).font = LABEL_FONT;
  ws.getCell(`B${r}`).value = (data.segue || {}).professional || '';
  r += 2;

  // 2. SCORECARD REVIEW
  r = writeSectionRow(ws, r, '2. SCORECARD REVIEW');
  r = writeHeaders(ws, r, ['Measurable / KPI', 'Owner', 'Goal', 'Actual', 'Status', 'Notes'], C);
  r = writeRows(ws, r, data.scorecardTable || [], C, { E: ['On Track', 'Off Track', 'At Risk'] });
  r++;

  // 3. OKR REVIEW
  r = writeSectionRow(ws, r, '3. OKR REVIEW');
  r = writeHeaders(ws, r, ['OKR / Rock Description', 'Owner', 'Due Date', 'Status', '% Done', 'Notes'], C);
  r = writeRows(ws, r, data.okrReviewTable || [], C, { D: ['On Track', 'Off Track', 'At Risk'] });
  r++;

  // 4. CUSTOMER / EMPLOYEE HEADLINES
  r = writeSectionRow(ws, r, '4. CUSTOMER / EMPLOYEE HEADLINES');
  r = writeHeaders(ws, r, ['Headline', 'Type', 'Reported By', 'Action Needed?', 'Add to IDS?', 'Notes'], C);
  r = writeRows(ws, r, data.headlinesTable || [], C, { B: ['Customer', 'Employee'], D: ['Yes', 'No'], E: ['Yes', 'No'] });
  r++;

  // 5. TO-DO LIST REVIEW
  r = writeSectionRow(ws, r, '5. TO-DO LIST REVIEW');
  r = writeHeaders(ws, r, ["Last Week's To-Do", 'Owner', 'Due Date', 'Status', 'Add to IDS?', 'Notes'], C);
  const todos = data.todoReviewTable || [];
  r = writeRows(ws, r, todos, C, { D: ['Open', 'Done', 'Carry Over'], E: ['Yes', 'No'] });
  let done = 0;
  todos.forEach((row: string[]) => { if (row[3] === 'Done') done++; });
  ws.getCell(`A${r}`).value = 'Completion:'; ws.getCell(`A${r}`).font = LABEL_FONT;
  ws.getCell(`B${r}`).value = `${done} / ${todos.length} done`;
  r += 2;

  // 6. IDS — IDENTIFY, DISCUSS, SOLVE
  r = writeSectionRow(ws, r, '6. IDS \u2014 IDENTIFY, DISCUSS, SOLVE');
  r = writeSubHeader(ws, r, 'Issues List');
  r = writeHeaders(ws, r, ['Issue / Obstacle', 'Raised By', 'Priority', 'Status', 'Time Est.', 'Next Mtg?'], C);
  r = writeRows(ws, r, data.issuesListTable || [], C, { C: ['High', 'Medium', 'Low'], D: ['Open', 'Solved', 'Next Meeting', 'Dropped'], F: ['Yes', 'No'] });
  r++;

  const idsBlocks = data.idsBlocks || [];
  idsBlocks.forEach((block: any, i: number) => {
    r = writeSubHeader(ws, r, `Issue #${i + 1}`);
    ws.getCell(`A${r}`).value = 'Issue:'; ws.getCell(`A${r}`).font = LABEL_FONT;
    ws.getCell(`B${r}`).value = (block.fields || [])[0] || '';
    r++;
    ws.getCell(`A${r}`).value = 'Root Cause:'; ws.getCell(`A${r}`).font = LABEL_FONT;
    ws.getCell(`B${r}`).value = (block.fields || [])[1] || '';
    r++;
    ws.getCell(`A${r}`).value = 'Solution:'; ws.getCell(`A${r}`).font = LABEL_FONT;
    ws.getCell(`B${r}`).value = (block.fields || [])[2] || '';
    r++;
    r = writeSubHeader(ws, r, 'New To-Do(s)');
    r = writeHeaders(ws, r, ['To-Do', 'Owner', 'Due Date', 'Priority', 'Status', 'Notes'], C);
    r = writeRows(ws, r, block.todos || [], C, { D: ['High', 'Medium', 'Low'], E: ['Not Started', 'In Progress', 'Done'] });
    r++;
  });
  r++;

  // 7. CONCLUDE
  r = writeSectionRow(ws, r, '7. CONCLUDE');

  r = writeSubHeader(ws, r, "New To-Do List \u2014 This Week's Commitments");
  r = writeHeaders(ws, r, ['To-Do', 'Owner', 'Due Date', 'Priority', 'Status', 'Notes'], C);
  r = writeRows(ws, r, data.newTodoTable || [], C, { D: ['High', 'Medium', 'Low'], E: ['Not Started', 'In Progress', 'Done'] });
  r++;

  r = writeSubHeader(ws, r, 'Cascading Messages \u2014 What needs to be shared?');
  r = writeHeaders(ws, r, ['Message', 'To Whom', 'By When', 'By Whom', 'Channel', 'Done?'], C);
  r = writeRows(ws, r, data.cascadingTable || [], C, { F: ['Yes', 'No'] });
  r++;

  r = writeSubHeader(ws, r, 'Meeting Rating \u2014 Rate 1-10');
  r = writeHeaders(ws, r, ['Team Member', 'Rating (1-10)', 'Quick Comment'], ['A', 'B', 'C']);
  const ratingTable = data.ratingTable || [];
  const ratingValues = data.ratingValues || [];
  let ratingSum = 0, ratingCount = 0;
  ratingTable.forEach((row: string[], i: number) => {
    ws.getCell(`A${r}`).value = row[0] || '';    ws.getCell(`A${r}`).border = BORDERS;
    const rv = parseInt(ratingValues[i]) || 0;
    ws.getCell(`B${r}`).value = rv > 0 ? rv : '';  ws.getCell(`B${r}`).border = BORDERS;
    ws.getCell(`C${r}`).value = row[2] || '';       ws.getCell(`C${r}`).border = BORDERS;
    if (rv > 0) { ratingSum += rv; ratingCount++; }
    r++;
  });
  ws.getCell(`A${r}`).value = 'Average Rating:'; ws.getCell(`A${r}`).font = LABEL_FONT;
  ws.getCell(`B${r}`).value = ratingCount > 0 ? parseFloat((ratingSum / ratingCount).toFixed(1)) : '';
}

function buildScorecardSheet(wb: ExcelJS.Workbook, data: Record<string, any>): void {
  const rows = data.scorecardFullTable as string[][] | undefined;
  if (!rows || rows.length === 0) return;

  const ws = wb.addWorksheet('Scorecard');
  ws.getColumn(1).width = 30; ws.getColumn(2).width = 18; ws.getColumn(3).width = 10;
  for (let i = 4; i <= 16; i++) ws.getColumn(i).width = 10;
  const cols = 'ABCDEFGHIJKLMNOP'.split('');

  let r = 1;
  ws.mergeCells('A1:P1');
  ws.getCell('A1').value = 'Scorecard Tracker (Rolling 13 Weeks)';
  ws.getCell('A1').font = { bold: true, size: 14 };
  r = 3;

  const headers = ['Measurable / KPI', 'Owner', 'Goal'];
  for (let w = 1; w <= 13; w++) headers.push(`Wk ${w}`);
  r = writeHeaders(ws, r, headers, cols);

  rows.forEach(row => {
    cols.forEach((col, ci) => {
      const cell = ws.getCell(`${col}${r}`);
      cell.value = row[ci] || '';
      cell.border = BORDERS;
    });
    r++;
  });
}

function buildOkrsSheet(wb: ExcelJS.Workbook, data: Record<string, any>): void {
  const rows = data.okrFullTable as string[][] | undefined;
  if (!rows || rows.length === 0) return;

  const ws = wb.addWorksheet('OKRs');
  [5, 35, 18, 12, 12, 10, 12, 25].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const cols = 'ABCDEFGH'.split('');

  let r = 1;
  ws.mergeCells('A1:H1');
  ws.getCell('A1').value = 'OKR Tracker (Rocks / 90-Day Priorities)';
  ws.getCell('A1').font = { bold: true, size: 14 };
  r = 3;

  r = writeHeaders(ws, r, ['#', 'OKR / Rock Description', 'Owner', 'Due Date', 'Priority', '% Done', 'Status', 'Notes'], cols);
  rows.forEach((row, i) => {
    ws.getCell(`A${r}`).value = i + 1; ws.getCell(`A${r}`).border = BORDERS;
    for (let ci = 0; ci < 7 && ci < row.length; ci++) {
      const cell = ws.getCell(`${cols[ci + 1]}${r}`);
      cell.value = row[ci] || '';
      cell.border = BORDERS;
    }
    r++;
  });

  const keyResults = data.keyResults as string[][][] | undefined;
  if (keyResults) {
    r += 2;
    keyResults.forEach((krRows, ki) => {
      if (krRows.length === 0) return;
      ws.mergeCells(`A${r}:H${r}`);
      ws.getCell(`A${r}`).value = `Key Results for OKR #${ki + 1}`;
      ws.getCell(`A${r}`).font = { bold: true, size: 11 };
      r++;
      r = writeHeaders(ws, r, ['#', 'Key Result', 'Owner', 'Due Date', 'Priority', '% Done', 'Status', 'Notes'], cols);
      krRows.forEach((row, ri) => {
        ws.getCell(`A${r}`).value = ri + 1; ws.getCell(`A${r}`).border = BORDERS;
        for (let ci = 0; ci < 7 && ci < row.length; ci++) {
          const cell = ws.getCell(`${cols[ci + 1]}${r}`);
          cell.value = row[ci] || '';
          cell.border = BORDERS;
        }
        r++;
      });
      r++;
    });
  }
}

// ── Read helpers (marker-based scanning for imported files) ──

function readWorkbookToJson(wb: ExcelJS.Workbook, ws: ExcelJS.Worksheet): Record<string, any> {
  const c = (ref: string) => cellStr(ws, ref);
  const rowCount = ws.rowCount;
  const C6 = ['A', 'B', 'C', 'D', 'E', 'F'];

  // Scan all rows for section markers
  const sec: Record<string, number> = {};
  const idsBlockRows: number[] = [];

  for (let r = 1; r <= rowCount; r++) {
    const a = c(`A${r}`);
    const au = a.toUpperCase();
    if ((au.startsWith('TEAM:') || au === 'TEAM') && !sec.meta) sec.meta = r;
    if (au.includes('SEGUE') && !sec.segue) sec.segue = r;
    if (au.includes('SCORECARD REVIEW') && !sec.scorecard) sec.scorecard = r;
    if (au.includes('OKR REVIEW') && !sec.okrReview) sec.okrReview = r;
    if (au.includes('HEADLINE') && !sec.headlines) sec.headlines = r;
    if ((au.includes('TO-DO LIST REVIEW') || au.includes('TODO LIST REVIEW')) && !sec.todoReview) sec.todoReview = r;
    if (au.includes('IDS') && (au.includes('IDENTIFY') || au.includes('DISCUSS')) && !sec.ids) sec.ids = r;
    if (a === 'Issue:') idsBlockRows.push(r);
    if (au.includes('CONCLUDE') && !sec.conclude) sec.conclude = r;
    if ((au.includes('NEW TO-DO LIST') || au.includes('COMMITMENTS')) && sec.conclude && !sec.newTodo) sec.newTodo = r;
    if (au.includes('CASCADING') && !sec.cascading) sec.cascading = r;
    if (au.includes('MEETING RATING') && !sec.rating) sec.rating = r;
    if (au.includes('AVERAGE RATING') || au.includes('AVG RATING')) sec.avgRating = r;
  }

  // Ordered boundaries for finding section ends
  const allBounds = Object.values(sec).concat(idsBlockRows.map(r => r - 1)).sort((a, b) => a - b);
  const nextAfter = (row: number): number => {
    for (const b of allBounds) { if (b > row + 1) return b; }
    return rowCount + 1;
  };

  // Read non-empty rows between two row numbers
  function readTable(start: number, end: number): string[][] {
    const rows: string[][] = [];
    for (let r = start; r < end; r++) {
      const row = C6.map(col => c(`${col}${r}`));
      if (row.every(v => !v)) continue;
      // Skip rows that look like headers/labels
      const a0 = row[0].toUpperCase();
      if (a0.includes('COMPLETION:') || a0.includes('AVERAGE RATING')) continue;
      rows.push(row);
    }
    return rows;
  }

  const data: Record<string, any> = {};

  // Meta
  if (sec.meta) {
    const mr = sec.meta;
    data.meta = {
      team: c(`B${mr}`), date: c(`E${mr}`),
      facilitator: c(`B${mr + 1}`), scribe: c(`E${mr + 1}`),
      start: c(`B${mr + 2}`), end: c(`E${mr + 2}`),
    };
  } else {
    data.meta = {};
  }

  // Segue
  if (sec.segue) {
    let personal = '', professional = '';
    const end = nextAfter(sec.segue);
    for (let r = sec.segue + 1; r < Math.min(sec.segue + 5, end); r++) {
      const au = c(`A${r}`).toUpperCase();
      if (au.includes('PERSONAL')) personal = c(`B${r}`);
      if (au.includes('PROFESSIONAL')) professional = c(`B${r}`);
    }
    data.segue = { personal, professional };
  }

  // Standard table sections: section header row + table header row = data starts at +2
  if (sec.scorecard) data.scorecardTable = readTable(sec.scorecard + 2, nextAfter(sec.scorecard));
  if (sec.okrReview) data.okrReviewTable = readTable(sec.okrReview + 2, nextAfter(sec.okrReview));
  if (sec.headlines) data.headlinesTable = readTable(sec.headlines + 2, nextAfter(sec.headlines));
  if (sec.todoReview) data.todoReviewTable = readTable(sec.todoReview + 2, nextAfter(sec.todoReview));

  // Issues list: IDS section header + "Issues List" sub-header + table header = +3
  if (sec.ids) {
    const end = idsBlockRows.length > 0 ? idsBlockRows[0] - 1 : nextAfter(sec.ids);
    data.issuesListTable = readTable(sec.ids + 3, end);
  }

  // IDS detail blocks
  const idsBlocks: any[] = [];
  idsBlockRows.forEach((blockRow, bi) => {
    const fields = [c(`B${blockRow}`), c(`B${blockRow + 1}`), c(`B${blockRow + 2}`)];
    const isPlaceholder = (s: string) => !s || s.startsWith('Describe the real') || s.startsWith("Ask 'why?'") || s.startsWith('Agreed solution');
    if (isPlaceholder(fields[0]) && isPlaceholder(fields[1]) && isPlaceholder(fields[2])) return;
    // After Issue/Root Cause/Solution: sub-header(+3), table header(+4), data(+5)
    const todoStart = blockRow + 5;
    const todoEnd = bi + 1 < idsBlockRows.length ? idsBlockRows[bi + 1] - 1 : (sec.conclude || nextAfter(blockRow));
    const todos = readTable(todoStart, todoEnd);
    idsBlocks.push({ fields, todos });
  });
  data.idsBlocks = idsBlocks;

  // Conclude sub-sections
  if (sec.newTodo) {
    const end = sec.cascading || sec.rating || nextAfter(sec.newTodo);
    data.newTodoTable = readTable(sec.newTodo + 2, end);
  } else if (sec.conclude) {
    const end = sec.cascading || sec.rating || nextAfter(sec.conclude);
    data.newTodoTable = readTable(sec.conclude + 3, end);
  }

  if (sec.cascading) {
    const end = sec.rating || nextAfter(sec.cascading);
    data.cascadingTable = readTable(sec.cascading + 2, end);
  }

  if (sec.rating) {
    const end = sec.avgRating || nextAfter(sec.rating);
    const ratingTable: string[][] = [];
    const ratingValues: string[] = [];
    for (let r = sec.rating + 2; r < end; r++) {
      const name = c(`A${r}`);
      if (!name || name.toUpperCase().includes('AVERAGE')) break;
      ratingTable.push([name, '', c(`C${r}`)]);
      ratingValues.push(c(`B${r}`) || '0');
    }
    data.ratingTable = ratingTable;
    data.ratingValues = ratingValues;
  }

  // Read from Scorecard sheet
  const scSheet = wb.getWorksheet('Scorecard');
  if (scSheet) {
    const scC = (ref: string) => cellStr(scSheet, ref);
    let headerRow = 0;
    for (let r = 1; r <= scSheet.rowCount; r++) {
      if (scC(`A${r}`).toUpperCase().includes('MEASURABLE') || scC(`A${r}`).toUpperCase().includes('KPI')) { headerRow = r; break; }
    }
    if (headerRow) {
      const hasGoal = scC(`C${headerRow}`).toUpperCase().includes('GOAL');
      const scRows: string[][] = [];
      for (let r = headerRow + 1; r <= scSheet.rowCount; r++) {
        const name = scC(`A${r}`);
        if (!name) continue;
        if (hasGoal) {
          const row = [name, scC(`B${r}`), scC(`C${r}`)];
          for (let w = 0; w < 13; w++) row.push(scC(`${String.fromCharCode(68 + w)}${r}`)); // D-P
          scRows.push(row);
        } else {
          const row = [name, scC(`B${r}`), ''];
          for (let w = 0; w < 13; w++) row.push(scC(`${String.fromCharCode(67 + w)}${r}`)); // C-O
          scRows.push(row);
        }
      }
      if (scRows.length > 0) data.scorecardFullTable = scRows;
      if (!data.scorecardTable || data.scorecardTable.length === 0 ||
          data.scorecardTable.every((r: string[]) => !r[0] || r[0] === '[object Object]')) {
        data.scorecardTable = scRows.map((r: string[]) => [r[0], r[1], '', '', '', '']);
      }
    }
  }

  // Read from OKRs sheet
  const okrSheet = wb.getWorksheet('OKRs');
  if (okrSheet) {
    const okC = (ref: string) => cellStr(okrSheet, ref);
    let headerRow = 0;
    for (let r = 1; r <= okrSheet.rowCount; r++) {
      const b = okC(`B${r}`).toUpperCase();
      if (b.includes('OKR') || b.includes('DESCRIPTION') || b.includes('ROCK')) { headerRow = r; break; }
    }
    if (headerRow) {
      const okrRows: string[][] = [];
      for (let r = headerRow + 1; r <= okrSheet.rowCount; r++) {
        const desc = okC(`B${r}`);
        if (!desc) {
          if (okC(`A${r}`).toUpperCase().includes('KEY RESULT')) break;
          continue;
        }
        okrRows.push([desc, okC(`C${r}`), okC(`D${r}`), okC(`E${r}`), okC(`F${r}`), okC(`G${r}`), okC(`H${r}`)]);
      }
      if (okrRows.length > 0) data.okrFullTable = okrRows;
      if (!data.okrReviewTable || data.okrReviewTable.length === 0 ||
          data.okrReviewTable.every((r: string[]) => !r[0] || r[0] === '[object Object]')) {
        data.okrReviewTable = okrRows.map((r: string[]) => [r[0], r[1], r[2], r[5], r[4], r[6]]);
      }

      // Key results
      const keyResults: string[][][] = [];
      let currentKR: string[][] = [];
      let inKR = false;
      for (let r = headerRow + okrRows.length + 1; r <= okrSheet.rowCount; r++) {
        const au = okC(`A${r}`).toUpperCase();
        if (au.includes('KEY RESULT')) {
          if (inKR && currentKR.length > 0) keyResults.push(currentKR);
          currentKR = [];
          inKR = true;
          r++; // skip table header row
          continue;
        }
        if (inKR) {
          const kr = okC(`B${r}`);
          if (!kr) { if (currentKR.length > 0) { keyResults.push(currentKR); currentKR = []; inKR = false; } continue; }
          currentKR.push([kr, okC(`C${r}`), okC(`D${r}`), okC(`E${r}`), okC(`F${r}`), okC(`G${r}`), okC(`H${r}`)]);
        }
      }
      if (inKR && currentKR.length > 0) keyResults.push(currentKR);
      if (keyResults.some(kr => kr.length > 0)) data.keyResults = keyResults;
    }
  }

  return data;
}
