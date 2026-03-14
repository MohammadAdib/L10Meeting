import express from 'express';
import path from 'path';
import fs from 'fs';
import { exec } from 'child_process';
import ExcelJS from 'exceljs';

const isPackaged = !!(process as any).pkg;
const exeDir = isPackaged ? path.dirname(process.execPath) : path.join(__dirname, '..');
const staticDir = path.join(__dirname, '..', 'dist');
const dataDir = path.join(exeDir, 'data');
const templatePath = path.join(exeDir, 'L10_Meeting_Template.xlsx');

// Ensure data directory exists
fs.mkdirSync(dataDir, { recursive: true });
fs.mkdirSync(path.join(dataDir, 'Departments'), { recursive: true });

const app = express();

// Import XLSX — must be registered before global JSON parser to preserve raw body
app.post('/api/departments/:name/meetings/import', express.raw({ type: '*/*', limit: '50mb' }), (req, res) => {
  try {
    if (!req.body || !Buffer.isBuffer(req.body) || req.body.length === 0) {
      res.status(400).json({ error: 'No file data received' }); return;
    }
    const deptName = req.params.name;
    const dir = path.join(deptDir(deptName), 'meetings');
    fs.mkdirSync(dir, { recursive: true });

    const today = new Date().toISOString().split('T')[0];
    let baseName = `L10_${deptName}_${today}`;
    let fileName = `${baseName}.xlsx`;
    let suffix = 1;
    while (fs.existsSync(path.join(dir, fileName))) {
      fileName = `${baseName}-${suffix}.xlsx`;
      suffix++;
    }
    const id = fileName.replace('.xlsx', '');
    fs.writeFileSync(path.join(dir, fileName), req.body);
    res.json({ id });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

app.use(express.json({ limit: '10mb' }));

// Serve static files
app.use(express.static(staticDir));

// Helper: get department directory path (with validation)
function deptDir(name: string): string {
  const dir = path.join(dataDir, 'Departments', name);
  if (!dir.startsWith(path.join(dataDir, 'Departments'))) throw new Error('Forbidden');
  return dir;
}

// ── Excel helpers ──

/** Write meeting JSON data into the visible "L10 Meeting" worksheet */
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

  // Scorecard Review (rows 14-20, cols A-F)
  writeTable(ws, data.scorecardTable, 14, 7, ['A', 'B', 'C', 'D', 'E', 'F']);

  // OKR Review (rows 26-31, cols A-F)
  writeTable(ws, data.okrReviewTable, 26, 6, ['A', 'B', 'C', 'D', 'E', 'F']);

  // Headlines (rows 37-42, cols A-F)
  writeTable(ws, data.headlinesTable, 37, 6, ['A', 'B', 'C', 'D', 'E', 'F']);

  // To-Do Review (rows 47-53, cols A-F)
  writeTable(ws, data.todoReviewTable, 47, 7, ['A', 'B', 'C', 'D', 'E', 'F']);

  // Completion rate
  const todos = data.todoReviewTable || [];
  let done = 0;
  todos.forEach((r: string[]) => { if (r[3] === 'Done') done++; });
  ws.getCell('E54').value = `${done} / ${todos.length} done`;

  // Issues List (rows 60-75, cols A-F)
  writeTable(ws, data.issuesListTable, 60, 16, ['A', 'B', 'C', 'D', 'E', 'F']);

  // IDS Issue Detail Blocks
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

  // New To-Dos (rows 171-181, cols A-F)
  writeTable(ws, data.newTodoTable, 171, 11, ['A', 'B', 'C', 'D', 'E', 'F']);

  // Cascading Messages (rows 184-189, cols A-F)
  writeTable(ws, data.cascadingTable, 184, 6, ['A', 'B', 'C', 'D', 'E', 'F']);

  // Meeting Rating (rows 192-197, cols A-C)
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

/** Extract a cell value as a plain string, handling formulas and dates */
function cellStr(ws: ExcelJS.Worksheet, ref: string): string {
  const v = ws.getCell(ref).value;
  if (v === null || v === undefined) return '';
  // Formula objects: { formula, result }
  if (typeof v === 'object' && 'result' in (v as any)) {
    const r = (v as any).result;
    if (r === null || r === undefined) return '';
    if (r instanceof Date) return formatDateCell(r);
    return String(r);
  }
  // Date objects
  if (v instanceof Date) return formatDateCell(v);
  return String(v);
}

function formatDateCell(d: Date): string {
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/** Read meeting data from the "L10 Meeting" worksheet into JSON */
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

  // Detect IDS section offset: check if "Issue:" label is at A78 (our format) or A79 (external +1)
  const idsOffset = c('A79').startsWith('Issue:') ? 1 : 0;

  // IDS blocks (9 rows apart)
  const idsBlocks: any[] = [];
  for (let bi = 0; bi < 10; bi++) {
    const base = 77 + idsOffset + bi * 9;
    // Fields are at base+1 (Issue:), base+2 (Root Cause:), base+3 (Solution:)
    const fields = [c(`B${base + 1}`), c(`B${base + 2}`), c(`B${base + 3}`)];
    // Skip placeholder text
    const isPlaceholder = (s: string) => !s || s.startsWith('Describe the real') || s.startsWith("Ask 'why?'") || s.startsWith('Agreed solution');
    if (isPlaceholder(fields[0]) && isPlaceholder(fields[1]) && isPlaceholder(fields[2])) continue;
    // Todos start at base+4 (our format) or base+5 (external, has header row)
    const todoStart = idsOffset ? base + 5 : base + 4;
    idsBlocks.push({ fields, todos: readTable(ws, todoStart, 5, ['A', 'B', 'C', 'D', 'E', 'F']) });
  }
  data.idsBlocks = idsBlocks;

  // Conclude section (offset by same amount)
  const newTodoStart = 171 + idsOffset;
  const cascadingStart = 184 + idsOffset;
  const ratingStart = 192 + idsOffset;

  data.newTodoTable = readTable(ws, newTodoStart, 11, ['A', 'B', 'C', 'D', 'E', 'F']);
  data.cascadingTable = readTable(ws, cascadingStart, 6, ['A', 'B', 'C', 'D', 'E', 'F']);

  // Rating
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
    // Skip empty rows and header rows (contain column labels like "To-Do", "Owner", etc.)
    if (row.every(v => !v)) continue;
    rows.push(row);
  }
  return rows;
}

// ── Department APIs ──

// List departments
app.get('/api/departments', (_req, res) => {
  try {
    const depsDir = path.join(dataDir, 'Departments');
    if (!fs.existsSync(depsDir)) { res.json([]); return; }
    const dirs = fs.readdirSync(depsDir, { withFileTypes: true })
      .filter(d => d.isDirectory())
      .map(d => {
        const peoplePath = path.join(depsDir, d.name, 'people.txt');
        let peopleCount = 0;
        if (fs.existsSync(peoplePath)) {
          const content = fs.readFileSync(peoplePath, 'utf-8').trim();
          peopleCount = content ? content.split('\n').filter(Boolean).length : 0;
        }
        return { name: d.name, peopleCount };
      });
    res.json(dirs);
  } catch {
    res.json([]);
  }
});

// Create department
app.post('/api/departments', (req, res) => {
  try {
    const { name } = req.body;
    if (!name || typeof name !== 'string') { res.status(400).json({ error: 'Name required' }); return; }
    const dir = deptDir(name.trim());
    if (fs.existsSync(dir)) { res.status(409).json({ error: 'Already exists' }); return; }
    fs.mkdirSync(dir, { recursive: true });
    fs.mkdirSync(path.join(dir, 'meetings'), { recursive: true });
    fs.writeFileSync(path.join(dir, 'people.txt'), '', 'utf-8');
    res.json({ ok: true, name: name.trim() });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Rename department
app.put('/api/departments/:name', (req, res) => {
  try {
    const oldDir = deptDir(req.params.name);
    const { name: newName } = req.body;
    if (!newName || typeof newName !== 'string') { res.status(400).json({ error: 'New name required' }); return; }
    if (!fs.existsSync(oldDir)) { res.status(404).json({ error: 'Not found' }); return; }
    const newDir = deptDir(newName.trim());
    if (fs.existsSync(newDir)) { res.status(409).json({ error: 'Target name already exists' }); return; }
    fs.renameSync(oldDir, newDir);
    res.json({ ok: true, name: newName.trim() });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Delete department
app.delete('/api/departments/:name', (req, res) => {
  try {
    const dir = deptDir(req.params.name);
    if (!fs.existsSync(dir)) { res.status(404).json({ error: 'Not found' }); return; }
    fs.rmSync(dir, { recursive: true, force: true });
    res.json({ ok: true });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// ── People APIs ──

// Get people
app.get('/api/departments/:name/people', (req, res) => {
  try {
    const dir = deptDir(req.params.name);
    const filePath = path.join(dir, 'people.txt');
    if (!fs.existsSync(filePath)) { res.json([]); return; }
    const content = fs.readFileSync(filePath, 'utf-8').trim();
    const people = content ? content.split('\n').map(s => s.trim()).filter(Boolean) : [];
    res.json(people);
  } catch {
    res.json([]);
  }
});

// Save people
app.put('/api/departments/:name/people', (req, res) => {
  try {
    const dir = deptDir(req.params.name);
    const { people } = req.body;
    if (!Array.isArray(people)) { res.status(400).json({ error: 'people array required' }); return; }
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, 'people.txt'), people.join('\n'), 'utf-8');
    res.json({ ok: true });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// ── Meeting APIs (Excel-based) ──

// List meetings
app.get('/api/departments/:name/meetings', async (req, res) => {
  try {
    const dir = path.join(deptDir(req.params.name), 'meetings');
    if (!fs.existsSync(dir)) { res.json([]); return; }
    const fileNames = fs.readdirSync(dir)
      .filter(f => f.endsWith('.xlsx') && !f.startsWith('~$'));

    const results = await Promise.all(fileNames.map(async f => {
      const filePath = path.join(dir, f);
      const stat = fs.statSync(filePath);
      const id = f.replace('.xlsx', '');
      const dateMatch = f.match(/(\d{4}-\d{2}-\d{2}(-\d+)?)/);
      const date = dateMatch ? dateMatch[1] : id;

      // Read avg rating from file
      let avgRating = 0;
      try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(filePath);
        const dataWs = wb.getWorksheet('_data');
        if (dataWs) {
          const raw = dataWs.getCell('A1').value;
          if (raw && typeof raw === 'string') {
            const data = JSON.parse(raw);
            const vals = (data.ratingValues || []) as string[];
            let sum = 0, count = 0;
            vals.forEach((v: string) => { const n = parseInt(v); if (n > 0) { sum += n; count++; } });
            if (count > 0) avgRating = sum / count;
          }
        }
        if (!avgRating) {
          // Fallback: read from worksheet
          const ws = wb.getWorksheet('L10 Meeting');
          if (ws) {
            const idsOffset = cellStr(ws, 'A79').startsWith('Issue:') ? 1 : 0;
            const ratingStart = 192 + idsOffset;
            let sum = 0, count = 0;
            for (let i = 0; i < 10; i++) {
              const v = cellStr(ws, `B${ratingStart + i}`);
              const n = parseInt(v);
              if (n > 0) { sum += n; count++; }
            }
            if (count > 0) avgRating = sum / count;
          }
        }
      } catch { /* skip rating read errors */ }

      return { id, date, lastSaved: stat.mtime.toISOString(), avgRating };
    }));

    results.sort((a, b) => b.date.localeCompare(a.date));
    res.json(results);
  } catch {
    res.json([]);
  }
});

// Create new meeting
app.post('/api/departments/:name/meetings', async (req, res) => {
  try {
    const deptName = req.params.name;
    const dir = path.join(deptDir(deptName), 'meetings');
    fs.mkdirSync(dir, { recursive: true });
    const today = new Date().toISOString().split('T')[0];
    let baseName = `L10_${deptName}_${today}`;
    let fileName = `${baseName}.xlsx`;
    let suffix = 1;
    while (fs.existsSync(path.join(dir, fileName))) {
      fileName = `${baseName}-${suffix}.xlsx`;
      suffix++;
    }
    const id = fileName.replace('.xlsx', '');

    const wb = new ExcelJS.Workbook();
    if (fs.existsSync(templatePath)) {
      await wb.xlsx.readFile(templatePath);
    } else {
      wb.addWorksheet('L10 Meeting');
    }

    const bodyData = req.body || {};
    bodyData.createdAt = new Date().toISOString();
    bodyData.lastSaved = new Date().toISOString();

    const ws = wb.getWorksheet('L10 Meeting');
    if (ws) writeJsonToWorkbook(ws, bodyData);

    // Store full JSON in _data sheet
    let dataWs = wb.getWorksheet('_data');
    if (!dataWs) dataWs = wb.addWorksheet('_data');
    dataWs.getCell('A1').value = JSON.stringify(bodyData);

    await wb.xlsx.writeFile(path.join(dir, fileName));
    res.json({ id });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Load meeting
app.get('/api/departments/:name/meetings/:id', async (req, res) => {
  try {
    const filePath = path.join(deptDir(req.params.name), 'meetings', `${req.params.id}.xlsx`);
    if (!fs.existsSync(filePath)) { res.status(404).json({ error: 'Not found' }); return; }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);

    // Try _data sheet first (lossless round-trip)
    const dataWs = wb.getWorksheet('_data');
    if (dataWs) {
      const raw = dataWs.getCell('A1').value;
      if (raw && typeof raw === 'string') {
        try {
          res.json(JSON.parse(raw));
          return;
        } catch { /* fall through to cell reading */ }
      }
    }

    // Fallback: read from L10 Meeting sheet
    const ws = wb.getWorksheet('L10 Meeting');
    if (!ws) { res.status(404).json({ error: 'Invalid file' }); return; }
    res.json(readWorkbookToJson(ws));
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Save meeting
app.put('/api/departments/:name/meetings/:id', async (req, res) => {
  try {
    const dir = path.join(deptDir(req.params.name), 'meetings');
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, `${req.params.id}.xlsx`);
    const data = { ...req.body, lastSaved: new Date().toISOString() };

    const wb = new ExcelJS.Workbook();
    if (fs.existsSync(filePath)) {
      await wb.xlsx.readFile(filePath);
    } else if (fs.existsSync(templatePath)) {
      await wb.xlsx.readFile(templatePath);
    } else {
      wb.addWorksheet('L10 Meeting');
    }

    const ws = wb.getWorksheet('L10 Meeting');
    if (ws) writeJsonToWorkbook(ws, data);

    // Store full JSON losslessly
    let dataWs = wb.getWorksheet('_data');
    if (!dataWs) dataWs = wb.addWorksheet('_data');
    dataWs.getCell('A1').value = JSON.stringify(data);

    // Atomic write
    const tmpPath = filePath + '.tmp.xlsx';
    await wb.xlsx.writeFile(tmpPath);
    fs.renameSync(tmpPath, filePath);
    res.json({ ok: true });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Delete meeting
app.delete('/api/departments/:name/meetings/:id', (req, res) => {
  try {
    const filePath = path.join(deptDir(req.params.name), 'meetings', `${req.params.id}.xlsx`);
    if (!fs.existsSync(filePath)) { res.status(404).json({ error: 'Not found' }); return; }
    fs.unlinkSync(filePath);
    res.json({ ok: true });
  } catch {
    res.status(404).json({ error: 'Not found' });
  }
});

// Open meeting in Excel
app.post('/api/departments/:name/meetings/:id/open', (req, res) => {
  try {
    const filePath = path.join(deptDir(req.params.name), 'meetings', `${req.params.id}.xlsx`);
    if (!fs.existsSync(filePath)) { res.status(404).json({ error: 'Not found' }); return; }
    exec(`start "" "${filePath}"`);
    res.json({ ok: true });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// SPA fallback
app.get('*', (_req, res) => {
  res.sendFile(path.join(staticDir, 'index.html'));
});

// Find available port
const PORT = 3847;
const server = app.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  console.log(`L10 Meeting Tool running at ${url}`);
  exec(`start ${url}`);
});

// Graceful shutdown
process.on('SIGINT', () => { server.close(); process.exit(0); });
process.on('SIGTERM', () => { server.close(); process.exit(0); });
