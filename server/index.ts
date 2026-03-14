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

const app = express();
app.use(express.json({ limit: '10mb' }));

// Serve static files
app.use(express.static(staticDir));

// API: List saved meetings
app.get('/api/meetings', (_req, res) => {
  try {
    const files = fs.readdirSync(dataDir)
      .filter(f => f.endsWith('.json'))
      .map(f => {
        const stat = fs.statSync(path.join(dataDir, f));
        return { filename: f, modified: stat.mtime.toISOString() };
      })
      .sort((a, b) => b.modified.localeCompare(a.modified));
    res.json(files);
  } catch {
    res.json([]);
  }
});

// API: Load a meeting
app.get('/api/meetings/:filename', (req, res) => {
  const filePath = path.join(dataDir, req.params.filename);
  if (!filePath.startsWith(dataDir)) return res.status(403).json({ error: 'Forbidden' });
  try {
    const data = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    res.json(data);
  } catch {
    res.status(404).json({ error: 'Not found' });
  }
});

// API: Save a meeting
app.post('/api/meetings/:filename', (req, res) => {
  const filePath = path.join(dataDir, req.params.filename);
  if (!filePath.startsWith(dataDir)) return res.status(403).json({ error: 'Forbidden' });
  try {
    // Atomic write: write to temp file then rename
    const tmpPath = filePath + '.tmp';
    fs.writeFileSync(tmpPath, JSON.stringify(req.body, null, 2), 'utf-8');
    fs.renameSync(tmpPath, filePath);
    res.json({ ok: true });
  } catch (err: any) {
    res.status(500).json({ error: err.message });
  }
});

// API: Delete a meeting
app.delete('/api/meetings/:filename', (req, res) => {
  const filePath = path.join(dataDir, req.params.filename);
  if (!filePath.startsWith(dataDir)) return res.status(403).json({ error: 'Forbidden' });
  try {
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
