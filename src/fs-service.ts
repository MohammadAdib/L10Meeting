import ExcelJS from 'exceljs';
import blankTemplateUrl from './blank.xlsx?url';

// Cached blank template buffer — loaded once from bundled asset
let _blankTemplateBuffer: ArrayBuffer | null = null;

async function getBlankTemplate(): Promise<ArrayBuffer> {
  if (_blankTemplateBuffer) return _blankTemplateBuffer;
  const resp = await fetch(blankTemplateUrl);
  _blankTemplateBuffer = await resp.arrayBuffer();
  return _blankTemplateBuffer;
}

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

  // Drop a website shortcut if it doesn't already exist
  try {
    const shortcutName = 'L10 Meeting Manager.html';
    await _rootHandle.getFileHandle(shortcutName);
  } catch {
    try {
      const fh = await _rootHandle.getFileHandle('L10 Meeting Manager.html', { create: true });
      const w = await fh.createWritable();
      await w.write('<html><head><meta http-equiv="refresh" content="0;url=https://mohammadadib.github.io/L10Meeting/"><link rel="icon" href="https://mohammadadib.github.io/L10Meeting/icon.svg" type="image/svg+xml"></head></html>');
      await w.close();
    } catch { /* silent — don't block app init */ }
  }
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

    const ws = wb.getWorksheet('L10 Meeting');
    if (ws) return readWorkbookToJson(wb, ws);
    return null;
  } catch {
    return null;
  }
}

async function writeExcelData(dirHandle: FileSystemDirectoryHandle, fileName: string, data: Record<string, any>): Promise<void> {
  const wb = new ExcelJS.Workbook();

  // Try to read existing file as base; fall back to blank.xlsx template
  let loaded = false;
  try {
    const fileHandle = await dirHandle.getFileHandle(fileName);
    const file = await fileHandle.getFile();
    const buffer = await file.arrayBuffer();
    await wb.xlsx.load(buffer);
    loaded = true;
  } catch { /* file doesn't exist yet */ }

  if (!loaded) {
    // Use blank.xlsx as the base template — preserves all formatting, formulas, merged cells
    const templateBuffer = await getBlankTemplate();
    await wb.xlsx.load(templateBuffer);
  }

  const ws = wb.getWorksheet('L10 Meeting');
  if (ws) writeJsonToWorkbook(wb, ws, data);

  // Remove legacy _data sheet if present
  const dataWs = wb.getWorksheet('_data');
  if (dataWs) wb.removeWorksheet(dataWs.id);

  const buffer = await wb.xlsx.writeBuffer();
  const fileHandle = await dirHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(buffer);
  await writable.close();
}

/** Export meeting data to an Excel buffer (no folder needed) */
export async function exportMeetingToBuffer(data: Record<string, any>): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const templateBuffer = await getBlankTemplate();
  await wb.xlsx.load(templateBuffer);
  const ws = wb.getWorksheet('L10 Meeting');
  if (ws) writeJsonToWorkbook(wb, ws, data);
  return wb.xlsx.writeBuffer() as Promise<ArrayBuffer>;
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

export async function getMeetings(deptName: string): Promise<{ id: string; date: string; avgRating: number }[]> {
  const ck = cacheKey('meetings', deptName);
  const cached = cacheGet<{ id: string; date: string; avgRating: number }[]>(ck);
  if (cached) return cached;
  if (!_rootHandle) return [];
  try {
    const meetings = await getMeetingsHandle(deptName);
    const results: { id: string; date: string; avgRating: number }[] = [];

    for await (const entry of (meetings as any).values()) {
      if (entry.kind !== 'file' || !entry.name.endsWith('.xlsx') || entry.name.startsWith('~$')) continue;
      const id = entry.name.replace('.xlsx', '');
      const dateMatch = entry.name.match(/(\d{4}-\d{2}-\d{2}(-\d+)?)/);
      const date = dateMatch ? dateMatch[1] : id;
      results.push({ id, date, avgRating: 0 });
    }

    results.sort((a, b) => b.date.localeCompare(a.date));
    cacheSet(ck, results);
    return results;
  } catch {
    return [];
  }
}

/** Load ratings for meetings in background (expensive — parses each Excel file) */
export async function loadMeetingRatings(deptName: string): Promise<{ id: string; date: string; avgRating: number }[]> {
  const ck = cacheKey('meetings', deptName);
  if (!_rootHandle) return cacheGet(ck) ?? [];
  try {
    const meetings = await getMeetingsHandle(deptName);
    const results: { id: string; date: string; avgRating: number }[] = [];

    for await (const entry of (meetings as any).values()) {
      if (entry.kind !== 'file' || !entry.name.endsWith('.xlsx') || entry.name.startsWith('~$')) continue;
      const id = entry.name.replace('.xlsx', '');
      const dateMatch = entry.name.match(/(\d{4}-\d{2}-\d{2}(-\d+)?)/);
      const date = dateMatch ? dateMatch[1] : id;

      let avgRating = 0;
      let actualDate = date;
      try {
        const file = await entry.getFile();
        const buffer = await file.arrayBuffer();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const ws = wb.getWorksheet('L10 Meeting');
        if (ws) {
          const dateVal = cellStr(ws, 'E2');
          if (dateVal) actualDate = dateVal;
          let sum = 0, count = 0;
          for (let r = 192; r <= 201; r++) {
            const v = ws.getCell(`B${r}`).value;
            const n = typeof v === 'number' ? v : parseInt(String(v || ''));
            if (n > 0) { sum += n; count++; }
          }
          if (count > 0) avgRating = sum / count;
        }
      } catch { /* skip */ }

      results.push({ id, date: actualDate, avgRating });
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


// ── Excel helpers (ported from server/index.ts) ──

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

function readWorkbookToJson(wb: ExcelJS.Workbook, ws: ExcelJS.Worksheet): Record<string, any> {
  const c = (ref: string) => cellStr(ws, ref);
  const data: Record<string, any> = {
    meta: {
      team: c('B2'), date: c('E2'), facilitator: c('B3'),
      scribe: c('E3'), start: c('B4'), end: c('E4'),
    },
    segue: { personal: c('B8'), professional: c('B9') },
  };

  // Read scorecard/OKR review from L10 Meeting sheet
  data.scorecardTable = readTable(ws, 14, 7, ['A', 'B', 'C', 'D', 'E', 'F']);
  data.okrReviewTable = readTable(ws, 26, 6, ['A', 'B', 'C', 'D', 'E', 'F']);

  // Read full Scorecard from separate sheet if it exists
  const scSheet = wb.getWorksheet('Scorecard');
  if (scSheet) {
    const scC = (ref: string) => cellStr(scSheet, ref);
    // Scorecard data: rows 4-13, cols A (name), B (owner), C-O (wk1-wk13) — but our table is A=name, B=owner, C=goal, D-P=weeks
    // Template: Row 3 is header, data at 4-13, cols A=name, B=owner, C..M = wk1-wk13 (no Goal column)
    const scRows: string[][] = [];
    for (let r = 4; r <= 13; r++) {
      const name = scC(`A${r}`);
      const owner = scC(`B${r}`);
      if (!name || name.startsWith('Measurable')) continue;
      const row = [name, owner, '']; // name, owner, goal (no goal in template)
      for (let w = 0; w < 13; w++) {
        row.push(scC(`${String.fromCharCode(67 + w)}${r}`)); // C through O
      }
      scRows.push(row);
    }
    if (scRows.length > 0) data.scorecardFullTable = scRows;

    // Also populate scorecardTable from scorecard sheet names/owners if L10 Meeting had formula refs
    if (data.scorecardTable.length === 0 || data.scorecardTable.every((r: string[]) => !r[0] || r[0] === '[object Object]')) {
      data.scorecardTable = scRows.map((r: string[]) => [r[0], r[1], '', '', '', '']); // name, owner, goal, actual, status, notes
    }
  }

  // Read full OKRs from separate sheet if it exists
  const okrSheet = wb.getWorksheet('OKRs');
  if (okrSheet) {
    const okC = (ref: string) => cellStr(okrSheet, ref);
    // OKR data: rows 5-14, cols A=#, B=desc, C=owner, D=due, E=priority, F=%done, G=status, H=notes
    // The # column is static in the UI, so exclude it from stored data
    const okrRows: string[][] = [];
    for (let r = 5; r <= 14; r++) {
      const desc = okC(`B${r}`);
      if (!desc) continue;
      okrRows.push([desc, okC(`C${r}`), okC(`D${r}`), okC(`E${r}`), okC(`F${r}`), okC(`G${r}`), okC(`H${r}`)]);
    }
    if (okrRows.length > 0) data.okrFullTable = okrRows;

    // Also populate okrReviewTable from OKRs sheet if L10 Meeting had formula refs
    if (data.okrReviewTable.length === 0 || data.okrReviewTable.every((r: string[]) => !r[0] || r[0] === '[object Object]')) {
      data.okrReviewTable = okrRows.map((r: string[]) => [r[0], r[1], r[2], r[5], r[4], r[6]]); // desc, owner, due, status, %done, notes
    }

    // Read key results blocks
    const keyResults: string[][][] = [];
    // OKR #1 key results at row 19, #2 at 25, #3 at 31
    const krStarts = [19, 25, 31];
    for (const start of krStarts) {
      const krRows: string[][] = [];
      for (let r = start; r <= start + 2; r++) {
        const kr = okC(`B${r}`);
        if (!kr) continue;
        krRows.push([kr, okC(`C${r}`), okC(`D${r}`), okC(`E${r}`), okC(`F${r}`), okC(`G${r}`), okC(`H${r}`)]);
      }
      keyResults.push(krRows);
    }
    if (keyResults.some(kr => kr.length > 0)) data.keyResults = keyResults;
  }
  data.headlinesTable = readTable(ws, 37, 7, ['A', 'B', 'C', 'D', 'E', 'F']);
  data.todoReviewTable = readTable(ws, 47, 7, ['A', 'B', 'C', 'D', 'E', 'F']);
  data.issuesListTable = readTable(ws, 60, 16, ['A', 'B', 'C', 'D', 'E', 'F']);

  // Dynamically find IDS issue blocks by scanning for "Issue:" in column A
  const idsBlocks: any[] = [];
  for (let r = 78; r <= 168; r++) {
    const a = c(`A${r}`);
    if (a === 'Issue:') {
      const fields = [c(`B${r}`), c(`B${r + 1}`), c(`B${r + 2}`)];
      const isPlaceholder = (s: string) => !s || s.startsWith('Describe the real') || s.startsWith("Ask 'why?'") || s.startsWith('Agreed solution');
      if (isPlaceholder(fields[0]) && isPlaceholder(fields[1]) && isPlaceholder(fields[2])) continue;
      // Todo header ("New To-Do(s)") is at r+3, data starts at r+4
      const todos = readTable(ws, r + 4, 4, ['A', 'B', 'C', 'D', 'E', 'F']);
      idsBlocks.push({ fields, todos });
    }
  }
  data.idsBlocks = idsBlocks;

  // Conclude section — data rows (row after column headers in blank.xlsx)
  const newTodoStart = 171;
  const cascadingStart = 184;
  const ratingStart = 192;

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

function clearTable(ws: ExcelJS.Worksheet, startRow: number, maxRows: number, cols: string[]): void {
  for (let i = 0; i < maxRows; i++) {
    const r = startRow + i;
    for (const col of cols) {
      ws.getCell(`${col}${r}`).value = '';
    }
  }
}

function writeJsonToWorkbook(wb: ExcelJS.Workbook, ws: ExcelJS.Worksheet, data: Record<string, any>): void {
  const cols = ['A', 'B', 'C', 'D', 'E', 'F'];

  // Clear all data regions before writing to avoid stale data from old saves
  clearTable(ws, 14, 7, cols);   // scorecardTable
  clearTable(ws, 26, 6, cols);   // okrReviewTable
  clearTable(ws, 37, 7, cols);   // headlinesTable
  clearTable(ws, 47, 7, cols);   // todoReviewTable
  clearTable(ws, 60, 16, cols);  // issuesListTable
  clearTable(ws, 171, 11, cols); // newTodoTable
  clearTable(ws, 184, 6, cols);  // cascadingTable
  for (let i = 0; i < 10; i++) { // ratingTable
    const r = 192 + i;
    ws.getCell(`A${r}`).value = '';
    ws.getCell(`B${r}`).value = '';
    ws.getCell(`C${r}`).value = '';
  }
  // Clear IDS blocks (fields + todos)
  const issueHeaderRows = [77, 86, 95, 104, 113, 122, 131, 140, 149, 158];
  for (const base of issueHeaderRows) {
    ws.getCell(`B${base + 1}`).value = '';  // Issue
    ws.getCell(`B${base + 2}`).value = '';  // Root Cause
    ws.getCell(`B${base + 3}`).value = '';  // Solution
    clearTable(ws, base + 5, 4, cols);      // Todos
  }

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
  writeTable(ws, data.headlinesTable, 37, 7, ['A', 'B', 'C', 'D', 'E', 'F']);
  writeTable(ws, data.todoReviewTable, 47, 7, ['A', 'B', 'C', 'D', 'E', 'F']);

  const todos = data.todoReviewTable || [];
  let done = 0;
  todos.forEach((r: string[]) => { if (r[3] === 'Done') done++; });
  ws.getCell('E54').value = `${done} / ${todos.length} done`;

  writeTable(ws, data.issuesListTable, 60, 16, ['A', 'B', 'C', 'D', 'E', 'F']);

  // IDS block layout: base=header("Issue #N"), Issue:(base+1), Root Cause:(base+2), Solution:(base+3), Todo header(base+4), todos(base+5..base+8)
  const issueStarts = [77, 86, 95, 104, 113, 122, 131, 140, 149, 158];
  const idsBlocks = data.idsBlocks || [];
  idsBlocks.forEach((block: any, bi: number) => {
    if (bi >= issueStarts.length) return;
    const base = issueStarts[bi];
    const fields = block.fields || [];
    // base is the header row ("Issue #N"), fields go into Issue:/Root Cause:/Solution: rows
    if (fields[0]) ws.getCell(`B${base + 1}`).value = fields[0];
    if (fields[1]) ws.getCell(`B${base + 2}`).value = fields[1];
    if (fields[2]) ws.getCell(`B${base + 3}`).value = fields[2];
    // Todos: base=header row, +1=Issue, +2=Root Cause, +3=Solution, +4=Todo header, +5=first data row
    writeTable(ws, block.todos, base + 5, 4, ['A', 'B', 'C', 'D', 'E', 'F']);
  });

  writeTable(ws, data.newTodoTable, 171, 11, ['A', 'B', 'C', 'D', 'E', 'F']);
  writeTable(ws, data.cascadingTable, 184, 6, ['A', 'B', 'C', 'D', 'E', 'F']);

  const ratingTable = data.ratingTable || [];
  const ratingValues = data.ratingValues || [];
  ratingTable.forEach((row: string[], i: number) => {
    if (i >= 10) return;
    const r = 192 + i;
    ws.getCell(`A${r}`).value = row[0] || '';
    const rv = parseInt(ratingValues[i]) || 0;
    ws.getCell(`B${r}`).value = rv > 0 ? rv : '';
    ws.getCell(`C${r}`).value = row[2] || '';
  });
  // Row 202 has AVERAGE formula in the template — update its range to cover all 10 slots
  ws.getCell('B202').value = { formula: 'IFERROR(AVERAGE(B192:B201),"")' };

  const validations: { col: string; startRow: number; count: number; options: string[] }[] = [
    { col: 'E', startRow: 14, count: 7, options: ['On Track', 'Off Track', 'At Risk'] },
    { col: 'D', startRow: 26, count: 6, options: ['On Track', 'Off Track', 'At Risk'] },
    { col: 'B', startRow: 37, count: 7, options: ['Customer', 'Employee'] },
    { col: 'D', startRow: 37, count: 7, options: ['Yes', 'No'] },
    { col: 'E', startRow: 37, count: 7, options: ['Yes', 'No'] },
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
    for (let i = 0; i < 4; i++) {
      ws.getCell(`D${base + 5 + i}`).dataValidation = {
        type: 'list', allowBlank: true, formulae: ['"High,Medium,Low"'],
      };
      ws.getCell(`E${base + 5 + i}`).dataValidation = {
        type: 'list', allowBlank: true, formulae: ['"Not Started,In Progress,Done"'],
      };
    }
  }

  // Write to Scorecard sheet
  const scFullRows = data.scorecardFullTable as string[][] | undefined;
  if (scFullRows) {
    let scSheet = wb.getWorksheet('Scorecard');
    if (!scSheet) scSheet = wb.addWorksheet('Scorecard');
    scFullRows.forEach((row, i) => {
      if (i >= 10) return;
      const r = 4 + i;
      scSheet!.getCell(`A${r}`).value = row[0] || ''; // name
      scSheet!.getCell(`B${r}`).value = row[1] || ''; // owner
      // row[2] is goal, skip (not in scorecard sheet)
      // Weeks data starts at row[3]
      for (let w = 0; w < 13; w++) {
        scSheet!.getCell(`${String.fromCharCode(67 + w)}${r}`).value = row[3 + w] || '';
      }
    });
  }

  // Write to OKRs sheet
  const okrFullRows = data.okrFullTable as string[][] | undefined;
  if (okrFullRows) {
    let okrSheet = wb.getWorksheet('OKRs');
    if (!okrSheet) okrSheet = wb.addWorksheet('OKRs');
    okrFullRows.forEach((row, i) => {
      if (i >= 10) return;
      const r = 5 + i;
      // row = [desc, owner, due, priority, %done, status, notes] (no # column)
      okrSheet!.getCell(`A${r}`).value = i + 1; // row number
      okrSheet!.getCell(`B${r}`).value = row[0] || '';
      okrSheet!.getCell(`C${r}`).value = row[1] || '';
      okrSheet!.getCell(`D${r}`).value = row[2] || '';
      okrSheet!.getCell(`E${r}`).value = row[3] || '';
      okrSheet!.getCell(`F${r}`).value = row[4] || '';
      okrSheet!.getCell(`G${r}`).value = row[5] || '';
      okrSheet!.getCell(`H${r}`).value = row[6] || '';
    });

    // Write key results
    const keyResults = data.keyResults as string[][][] | undefined;
    if (keyResults) {
      const krStarts = [19, 25, 31];
      keyResults.forEach((krRows, ki) => {
        if (ki >= krStarts.length) return;
        krRows.forEach((row, ri) => {
          if (ri >= 3) return;
          const r = krStarts[ki] + ri;
          okrSheet!.getCell(`A${r}`).value = ri + 1; // row number
          for (let ci = 0; ci < 7 && ci < row.length; ci++) {
            okrSheet!.getCell(`${String.fromCharCode(66 + ci)}${r}`).value = row[ci] || ''; // B onwards
          }
        });
      });
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
