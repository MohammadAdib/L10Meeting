import express from 'express';
import path from 'path';
import fs from 'fs';
import { exec } from 'child_process';

const isPackaged = !!(process as any).pkg;
const exeDir = isPackaged ? path.dirname(process.execPath) : path.join(__dirname, '..');
const staticDir = path.join(__dirname, '..', 'dist');
const dataDir = path.join(exeDir, 'data');

// Ensure data directory exists
fs.mkdirSync(dataDir, { recursive: true });
fs.mkdirSync(path.join(dataDir, 'departments'), { recursive: true });

const app = express();
app.use(express.json({ limit: '10mb' }));

// Serve static files
app.use(express.static(staticDir));

// Helper: get department directory path (with validation)
function deptDir(name: string): string {
  const dir = path.join(dataDir, 'departments', name);
  if (!dir.startsWith(path.join(dataDir, 'departments'))) throw new Error('Forbidden');
  return dir;
}

// ── Department APIs ──

// List departments
app.get('/api/departments', (_req, res) => {
  try {
    const depsDir = path.join(dataDir, 'departments');
    if (!fs.existsSync(depsDir)) { res.json([]); return; }
    const dirs = fs.readdirSync(depsDir, { withFileTypes: true })
      .filter(d => d.isDirectory())
      .map(d => {
        const meetingsDir = path.join(depsDir, d.name, 'meetings');
        let meetingCount = 0;
        if (fs.existsSync(meetingsDir)) {
          meetingCount = fs.readdirSync(meetingsDir).filter(f => f.endsWith('.json')).length;
        }
        return { name: d.name, meetingCount };
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

// ── Meeting APIs ──

// List meetings
app.get('/api/departments/:name/meetings', (req, res) => {
  try {
    const dir = path.join(deptDir(req.params.name), 'meetings');
    if (!fs.existsSync(dir)) { res.json([]); return; }
    const files = fs.readdirSync(dir)
      .filter(f => f.endsWith('.json'))
      .map(f => {
        const stat = fs.statSync(path.join(dir, f));
        const id = f.replace('.json', '');
        // Try to read lastSaved from file
        let lastSaved = stat.mtime.toISOString();
        try {
          const data = JSON.parse(fs.readFileSync(path.join(dir, f), 'utf-8'));
          if (data.lastSaved) lastSaved = data.lastSaved;
        } catch { /* use stat time */ }
        return { id, date: id, lastSaved };
      })
      .sort((a, b) => b.id.localeCompare(a.id));
    res.json(files);
  } catch {
    res.json([]);
  }
});

// Create new meeting
app.post('/api/departments/:name/meetings', (req, res) => {
  try {
    const dir = path.join(deptDir(req.params.name), 'meetings');
    fs.mkdirSync(dir, { recursive: true });
    const today = new Date().toISOString().split('T')[0];
    let id = today;
    let suffix = 2;
    while (fs.existsSync(path.join(dir, `${id}.json`))) {
      id = `${today}-${suffix}`;
      suffix++;
    }
    const now = new Date().toISOString();
    const data = { createdAt: now, lastSaved: now, ...(req.body || {}) };
    fs.writeFileSync(path.join(dir, `${id}.json`), JSON.stringify(data, null, 2), 'utf-8');
    res.json({ id });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Load meeting
app.get('/api/departments/:name/meetings/:id', (req, res) => {
  try {
    const filePath = path.join(deptDir(req.params.name), 'meetings', `${req.params.id}.json`);
    if (!fs.existsSync(filePath)) { res.status(404).json({ error: 'Not found' }); return; }
    const data = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    res.json(data);
  } catch {
    res.status(404).json({ error: 'Not found' });
  }
});

// Save meeting
app.put('/api/departments/:name/meetings/:id', (req, res) => {
  try {
    const dir = path.join(deptDir(req.params.name), 'meetings');
    fs.mkdirSync(dir, { recursive: true });
    const filePath = path.join(dir, `${req.params.id}.json`);
    const data = { ...req.body, lastSaved: new Date().toISOString() };
    const tmpPath = filePath + '.tmp';
    fs.writeFileSync(tmpPath, JSON.stringify(data, null, 2), 'utf-8');
    fs.renameSync(tmpPath, filePath);
    res.json({ ok: true });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// Delete meeting
app.delete('/api/departments/:name/meetings/:id', (req, res) => {
  try {
    const filePath = path.join(deptDir(req.params.name), 'meetings', `${req.params.id}.json`);
    if (!fs.existsSync(filePath)) { res.status(404).json({ error: 'Not found' }); return; }
    fs.unlinkSync(filePath);
    res.json({ ok: true });
  } catch {
    res.status(404).json({ error: 'Not found' });
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
