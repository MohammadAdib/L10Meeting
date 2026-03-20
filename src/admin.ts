import { getLogoUrl, handleLogoClick } from './logo';
import { confirmDialog, populateTableRows } from './utils';
import { buildScorecardContent, buildOkrsContent, buildIDSContent, buildTodosContent } from './html';
import { addScorecardFullRow, addOkrFullRow, buildKeyResultBlocks, addIssueRow, addIDSIssue, addIDSTodoRow, resetIdsIssueCount } from './tables';
import { loadScorecardOkrData } from './storage';
import * as fs from './fs-service';
import blankTemplateUrl from './blank.xlsx?url';
import { showSettingsMenu } from './settings';
import { DEFAULT_ROWS, MAX_ROWS } from './types';

let _selectedDept: string | null = null;
let _peopleSaveTimer: ReturnType<typeof setTimeout> | null = null;

function formatCellValue(v: string): string {
  if (!v) return '';
  const m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return `${parseInt(m[2])}/${m[3]}/${m[1]}`;
  return v;
}

/** Populate a read-only dept table by creating input rows from data */
function populateDeptTable(selector: string, rows: string[][]): void {
  const tb = document.querySelector(`${selector} tbody`);
  if (!tb) return;
  const colCount = document.querySelectorAll(`${selector} thead th`).length;
  const count = Math.max(rows.length, 1);
  for (let i = 0; i < count; i++) {
    const tr = document.createElement('tr');
    const cols = rows[i] || [];
    tr.innerHTML = Array.from({ length: colCount }, (_, ci) =>
      `<td><input value="${formatCellValue(cols[ci] || '').replace(/"/g, '&quot;')}" disabled></td>`
    ).join('');
    tb.appendChild(tr);
  }
}

function buildCalendarView(todoRows: string[][], cascadingRows: string[][]): void {
  const container = document.getElementById('deptCalendarView');
  if (!container) return;

  // Collect items with dates
  type CalItem = { label: string; owner: string; date: string; type: 'todo' | 'cascading'; done: boolean };
  const items: CalItem[] = [];

  for (const r of todoRows) {
    if (!r[0]?.trim() || !r[2]?.trim()) continue;
    items.push({ label: r[0], owner: r[1] || '', date: r[2], type: 'todo', done: (r[4] || '').toLowerCase().includes('done') });
  }
  for (const r of cascadingRows) {
    if (!r[0]?.trim() || !r[2]?.trim()) continue;
    items.push({ label: r[0], owner: r[3] || '', date: r[2], type: 'cascading', done: (r[5] || '').toLowerCase().includes('yes') });
  }

  if (items.length === 0) return;

  // Group by date
  const byDate = new Map<string, CalItem[]>();
  for (const item of items) {
    const key = item.date;
    if (!byDate.has(key)) byDate.set(key, []);
    byDate.get(key)!.push(item);
  }

  // Sort dates
  const sortedDates = [...byDate.keys()].sort();
  const now = new Date();
  const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;

  const formatDate = (d: string) => {
    const m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return d;
    const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const dt = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
    return `${days[dt.getDay()]} ${parseInt(m[2])}/${m[3]}`;
  };

  let html = '';
  for (const date of sortedDates) {
    const dateItems = byDate.get(date)!;
    const isPast = date < today;
    const isToday = date === today;
    html += `<div class="cal-day${isToday ? ' cal-today' : ''}${isPast ? ' cal-past' : ''}">
      <div class="cal-date">${formatDate(date)}${isToday ? ' <span class="cal-today-badge">Today</span>' : ''}</div>
      <div class="cal-items">
        ${dateItems.map(item => `<div class="cal-item cal-${item.type}${item.done ? ' cal-done' : ''}">
          <span class="cal-dot"></span>
          <span class="cal-label">${item.label}</span>
          <span class="cal-owner">${item.owner}</span>
        </div>`).join('')}
      </div>
    </div>`;
  }

  container.innerHTML = html;
}

function savePeopleDebounced(deptName: string): void {
  if (_peopleSaveTimer) clearTimeout(_peopleSaveTimer);
  _peopleSaveTimer = setTimeout(async () => {
    const inputs = document.querySelectorAll<HTMLInputElement>('#peopleList .people-input');
    const names = Array.from(inputs).map(i => i.value.trim()).filter(Boolean);
    await fs.savePeople(deptName, names);
  }, 1000);
}

export async function renderAdminPortal(selectedDept?: string): Promise<void> {
  _selectedDept = selectedDept || null;
  const app = document.getElementById('app')!;

  let departments: { name: string; peopleCount: number }[] = [];
  try {
    departments = await fs.getDepartments();
  } catch { /* empty */ }

  // Auto-select first department if none specified
  if (!_selectedDept && departments.length > 0) {
    _selectedDept = departments[0].name;
  }

  const sidebarItems = departments.map(d => `
    <a class="sidebar-item${d.name === _selectedDept ? ' active' : ''}" data-dept="${d.name}">
      <span class="sidebar-label">${d.name}</span>
      <span class="sidebar-meta">${d.peopleCount}</span>
    </a>
  `).join('');

  app.innerHTML = `
    <div class="top-bar-wrapper">
      <div class="top-bar">
        <div class="top-bar-left">
          ${getLogoUrl() ? `<img src="${getLogoUrl()}" class="top-bar-logo">` : `<button class="top-bar-logo-placeholder" id="btnAddLogo">+ Add Logo</button>`}
          <div class="top-bar-title">L10 Meeting Manager</div>
        </div>
        <div class="top-bar-actions" style="opacity:1;pointer-events:auto">
          <button class="settings-btn" id="btnSettings" title="Data folder">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
            </svg>
          </button>
        </div>
      </div>
    </div>
    <div class="admin-layout">
      <nav class="admin-sidebar">
        <div class="admin-sidebar-header">
          <span>Departments</span>
          <button class="admin-sidebar-add" id="btnAddDept" title="Add Department">+</button>
        </div>
        <div class="admin-sidebar-list">
          ${sidebarItems || '<div class="admin-sidebar-empty">No departments</div>'}
        </div>
        <div class="admin-sidebar-footer">
          <a class="admin-sidebar-link" id="btnDownloadTemplate">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Template
          </a>
          <a class="admin-sidebar-link" id="btnAbout">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>
            About
          </a>
        </div>
      </nav>
      <div class="admin-content" id="adminContent">
        ${_selectedDept ? '' : '<div class="admin-empty-content">Create a department to get started.</div>'}
      </div>
    </div>
    <div class="toast"></div>
  `;

  // Logo click -> back to landing page
  document.querySelector('.top-bar-logo')?.addEventListener('click', () => {
    location.hash = '#/';
  });

  // Add logo button
  document.getElementById('btnAddLogo')?.addEventListener('click', () => {
    handleLogoClick(() => renderAdminPortal(_selectedDept || undefined));
  });

  // Settings gear
  document.getElementById('btnSettings')?.addEventListener('click', (e) => {
    showSettingsMenu(e.currentTarget as HTMLElement);
  });

  // Template download
  document.getElementById('btnDownloadTemplate')?.addEventListener('click', async (e) => {
    e.preventDefault();
    const resp = await fetch(blankTemplateUrl);
    const blob = await resp.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'L10_Template.xlsx';
    a.click();
    URL.revokeObjectURL(a.href);
  });

  // About dialog
  document.getElementById('btnAbout')?.addEventListener('click', (e) => {
    e.preventDefault();
    showAboutDialog();
  });

  // Wire sidebar clicks
  app.querySelectorAll<HTMLElement>('.sidebar-item[data-dept]').forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const dept = item.dataset.dept!;
      if (dept === _selectedDept) return;
      // Update active state
      app.querySelectorAll('.sidebar-item').forEach(i => i.classList.remove('active'));
      item.classList.add('active');
      _selectedDept = dept;
      // Update URL without triggering the router (which would full-reload the page)
      history.replaceState(null, '', `#/dept/${encodeURIComponent(dept)}`);
      loadDepartmentContent(dept);
    });
  });

  // Add department button
  document.getElementById('btnAddDept')?.addEventListener('click', async () => {
    const name = prompt('Department name:');
    if (!name || !name.trim()) return;
    const result = await fs.createDepartment(name.trim());
    if (!result.ok) {
      alert(result.error || 'Failed to create department');
      return;
    }
    await renderAdminPortal(name.trim());
  });

  // Load selected department content
  if (_selectedDept) {
    await loadDepartmentContent(_selectedDept);
  }
}

function buildRatingHtml(avg: number): string {
  if (avg <= 0) return '<span class="meeting-rating-val" style="color:var(--text-muted)">No Rating</span>';
  const fullStars = Math.round(avg);
  return Array.from({ length: 10 }, (_, i) =>
    `<span class="meeting-star${i < fullStars ? ' active' : ''}">\u2605</span>`
  ).join('') + `<span class="meeting-rating-val">${avg.toFixed(1)}</span>`;
}

function buildMeetingItemsHtml(meetings: { id: string; date: string; avgRating: number }[]): string {
  return meetings.map(m => {
    let displayDate = m.date;
    const dp = m.date.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (dp) displayDate = `${parseInt(dp[2])}/${dp[3]}/${dp[1]}`;
    return `
      <div class="meeting-item" data-id="${m.id}">
        <div class="meeting-item-date">${displayDate}</div>
        <div class="meeting-item-rating" id="rating-${m.id}"></div>
        <button class="meeting-item-delete btn btn-outline-dark btn-sm" data-id="${m.id}" title="Delete meeting">&times;</button>
      </div>
    `;
  }).join('');
}

async function loadDepartmentContent(deptName: string): Promise<void> {
  const content = document.getElementById('adminContent');
  if (!content) return;

  // Trigger fade-slide animation on content swap
  content.style.animation = 'none';
  content.offsetHeight; // force reflow
  content.style.animation = 'fadeSlideRight .3s ease forwards';

  let people: string[] = [];
  let meetings: { id: string; date: string; avgRating: number }[] = [];

  try {
    [people, meetings] = await Promise.all([
      fs.getPeople(deptName),
      fs.getMeetings(deptName),
    ]);
  } catch { /* empty */ }

  const peopleItems = people.map((p, i) => `
    <div class="people-item" data-index="${i}">
      <input class="people-input" value="${p.replace(/"/g, '&quot;')}" placeholder="Name">
      <button class="people-remove" data-index="${i}">&times;</button>
    </div>
  `).join('');

  const meetingItems = buildMeetingItemsHtml(meetings);

  content.innerHTML = `
    <div class="dept-header">
      <h1>${deptName}</h1>
      <div class="dept-header-actions">
        <button class="btn btn-outline" id="btnRenameDept">Rename</button>
        <button class="btn btn-outline" id="btnDeleteDept" style="border-color:var(--red);color:var(--red);">Delete</button>
      </div>
    </div>

    <div class="dept-layout">
      <div class="dept-left">
        <div class="dept-section">
          <div class="dept-section-header">
            <h2>People</h2>
            <button class="btn btn-outline-dark" id="btnAddPerson">+ Add Person</button>
          </div>
          <div class="people-list" id="peopleList">
            ${peopleItems || '<div class="admin-empty" style="padding:12px;">No people added yet.</div>'}
          </div>
        </div>

        <div class="dept-section">
          <div class="dept-section-header">
            <h2>Meetings</h2>
            <div style="display:flex;gap:8px;">
              <button class="btn btn-outline" id="btnImportCsv">Import</button>
              <button class="btn btn-green" id="btnNewMeeting">+ New Meeting</button>
            </div>
            <input type="file" id="csvFileInput" accept=".xlsx" style="display:none">
          </div>
          <div class="meetings-list">
            ${meetingItems || '<div class="admin-empty" style="padding:12px;">No meetings yet.</div>'}
          </div>
          <canvas id="ratingsChart" class="ratings-chart" style="display:none"></canvas>
        </div>
      </div>

      <div class="dept-right dept-readonly${meetings.length === 0 ? ' dept-no-meetings' : ' dept-loading'}">
        <div class="dept-readonly-label">${meetings.length > 0 ? `From most recent meeting (${meetings[0].date})` : 'From most recent meeting — no meetings yet'}</div>
        <div class="dept-tabs">
          <button class="dept-tab active" data-dept-tab="todos">To-Dos</button>
          <button class="dept-tab" data-dept-tab="ids">IDS</button>
          <button class="dept-tab" data-dept-tab="scorecard">Scorecard</button>
          <button class="dept-tab" data-dept-tab="okrs">OKRs</button>
        </div>
        <div class="dept-tab-content active" id="dept-tab-todos">
          ${buildTodosContent()}
        </div>
        <div class="dept-tab-content" id="dept-tab-ids">
          ${buildIDSContent()}
        </div>
        <div class="dept-tab-content" id="dept-tab-scorecard">
          ${buildScorecardContent()}
        </div>
        <div class="dept-tab-content" id="dept-tab-okrs">
          ${buildOkrsContent()}
        </div>
      </div>
    </div>
  `;

  wireContentEvents(deptName);

  // ── Wire dept tab switching ──
  document.querySelectorAll<HTMLButtonElement>('.dept-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.dept-tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.dept-tab-content').forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(`dept-tab-${btn.dataset.deptTab}`)?.classList.add('active');
    });
  });

  // ── Populate scorecard/OKR/IDS ──
  if (meetings.length === 0) {
    // No meetings — show default empty rows (read-only)
    for (let i = 0; i < MAX_ROWS.scorecardFull; i++) addScorecardFullRow();
    for (let i = 1; i <= MAX_ROWS.okrFull; i++) addOkrFullRow('', i);
    buildKeyResultBlocks();
    const dr = document.querySelector('.dept-right');
    if (dr) {
      dr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea').forEach(el => {
        el.disabled = true;
      });
      dr.querySelectorAll<HTMLElement>('.person-picker').forEach(el => el.classList.add('disabled'));
      dr.querySelectorAll<HTMLElement>('.add-row-btn, .row-delete').forEach(el => {
        el.style.display = 'none';
      });
    }
  } else {
    fs.getMeetingData(deptName, meetings[0].id).then(lastMeetingData => {
      if (_selectedDept !== deptName) return;

      const scorecardRows = (lastMeetingData?.scorecardFullTable as string[][] | undefined) || [];
      const okrRows = (lastMeetingData?.okrFullTable as string[][] | undefined) || [];

      // Clear and rebuild with correct row counts
      const scTb = document.querySelector('#scorecardFullTable tbody');
      const okTb = document.querySelector('#okrFullTable tbody');
      const krContainer = document.getElementById('okrKeyResultsContainer');
      if (scTb) scTb.innerHTML = '';
      if (okTb) okTb.innerHTML = '';
      if (krContainer) krContainer.innerHTML = '';

      const scCount = Math.max(scorecardRows.length, MAX_ROWS.scorecardFull);
      for (let i = 0; i < scCount; i++) addScorecardFullRow();
      const okCount = Math.max(okrRows.length, MAX_ROWS.okrFull);
      for (let i = 1; i <= okCount; i++) addOkrFullRow('', i);
      buildKeyResultBlocks();

      if (lastMeetingData) {
        loadScorecardOkrData(lastMeetingData as Record<string, unknown>);
      }

      // ── Populate IDS data ──
      resetIdsIssueCount();
      const issuesRows = (lastMeetingData?.issuesListTable as string[][] | undefined) || [];
      const idsBlocks = (lastMeetingData?.idsBlocks as { fields: string[]; todos: string[][] }[] | undefined) || [];

      // Issues list table
      const issTb = document.querySelector('#issuesListTable tbody');
      if (issTb) {
        issTb.innerHTML = '';
        for (let i = 0; i < Math.max(issuesRows.length, 1); i++) addIssueRow();
        if (issuesRows.length > 0) populateTableRows('#issuesListTable', issuesRows);
      }

      // IDS detail blocks
      const idsContainer = document.getElementById('idsIssuesContainer');
      if (idsContainer) {
        idsContainer.innerHTML = '';
        for (let i = 0; i < idsBlocks.length; i++) addIDSIssue();
        const blocks = idsContainer.querySelectorAll('.ids-issue');
        idsBlocks.forEach((block, bi) => {
          if (bi >= blocks.length) return;
          const tas = blocks[bi].querySelectorAll<HTMLTextAreaElement>('.ids-field textarea');
          block.fields.forEach((v, fi) => { if (fi < tas.length) tas[fi].value = v; });
          const todos = block.todos;
          if (todos && todos.length > 0) {
            const issueN = bi + 1;
            const existingTodos = document.querySelectorAll(`#idsTodo-${issueN} tbody tr`).length;
            for (let i = existingTodos; i < todos.length; i++) addIDSTodoRow(issueN);
            populateTableRows(`#idsTodo-${issueN}`, todos);
          }
        });
      }

      // New To-Do and Cascading tables
      const newTodoRows = (lastMeetingData?.newTodoTable as string[][] | undefined) || [];
      const cascadingRows = (lastMeetingData?.cascadingTable as string[][] | undefined) || [];

      populateDeptTable('#deptNewTodoTable', newTodoRows);
      populateDeptTable('#deptCascadingTable', cascadingRows);
      buildCalendarView(newTodoRows, cascadingRows);

      // Apply read-only and reveal
      const dr = document.querySelector<HTMLElement>('.dept-right');
      if (dr) {
        dr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea').forEach(el => {
          el.disabled = true;
        });
        dr.querySelectorAll<HTMLElement>('.person-picker').forEach(el => el.classList.add('disabled'));
        dr.querySelectorAll<HTMLElement>('.add-row-btn, .row-delete').forEach(el => {
          el.style.display = 'none';
        });
        dr.classList.remove('dept-loading');
      }
    });

    // ── Background: load meeting ratings + actual dates ──
    fs.loadMeetingRatings(deptName).then(rated => {
      if (_selectedDept !== deptName) return;
      for (const m of rated) {
        const el = document.getElementById(`rating-${m.id}`);
        if (el) {
          el.style.opacity = '0';
          el.innerHTML = buildRatingHtml(m.avgRating);
          requestAnimationFrame(() => { el.style.transition = 'opacity 150ms ease'; el.style.opacity = '1'; });
        }
        // Update date from file contents
        const item = document.querySelector<HTMLElement>(`.meeting-item[data-id="${m.id}"]`);
        const dateEl = item?.querySelector('.meeting-item-date');
        if (dateEl && m.date) {
          const dp = m.date.match(/^(\d{4})-(\d{2})-(\d{2})/);
          if (dp) dateEl.textContent = `${parseInt(dp[2])}/${dp[3]}/${dp[1]}`;
        }
      }
      // Draw ratings chart
      const withRatings = rated.filter(m => m.avgRating > 0).reverse().slice(-12); // last 12, chronological
      drawRatingsChart(withRatings);
    });
  }
}

function drawRatingsChart(meetings: { date: string; avgRating: number }[]): void {
  const canvas = document.getElementById('ratingsChart') as HTMLCanvasElement | null;
  if (!canvas || meetings.length < 2) return;
  canvas.style.display = '';

  function render() {
    const dpr = window.devicePixelRatio || 1;
    const w = canvas!.parentElement!.clientWidth - 40; // account for section padding
    const h = 100;
    canvas!.width = w * dpr;
    canvas!.height = h * dpr;
    canvas!.style.width = `${w}px`;
    canvas!.style.height = `${h}px`;

    const ctx = canvas!.getContext('2d')!;
    ctx.scale(dpr, dpr);

    const maxR = 10;
    const padL = 24, padR = 8, padT = 12, padB = 20;
    const chartW = w - padL - padR;
    const chartH = h - padT - padB;

    ctx.fillStyle = 'rgba(255,255,255,.3)';
    ctx.font = '10px system-ui';
    ctx.textAlign = 'right';
    for (const v of [0, 5, 10]) {
      const y = padT + chartH - (v / maxR) * chartH;
      ctx.fillText(String(v), padL - 6, y + 3);
      ctx.strokeStyle = 'rgba(255,255,255,.06)';
      ctx.beginPath();
      ctx.moveTo(padL, y);
      ctx.lineTo(padL + chartW, y);
      ctx.stroke();
    }

    ctx.textAlign = 'center';
    ctx.fillStyle = 'rgba(255,255,255,.3)';
    const step = Math.max(1, Math.floor(meetings.length / 5));
    meetings.forEach((m, i) => {
      if (i % step !== 0 && i !== meetings.length - 1) return;
      const x = padL + (i / (meetings.length - 1)) * chartW;
      const dp = m.date.match(/^(\d{4})-(\d{2})-(\d{2})/);
      const label = dp ? `${parseInt(dp[2])}/${dp[3]}` : '';
      ctx.fillText(label, x, h - 4);
    });

    ctx.beginPath();
    ctx.strokeStyle = '#ff8c42';
    ctx.lineWidth = 2;
    ctx.lineJoin = 'round';
    meetings.forEach((m, i) => {
      const x = padL + (i / (meetings.length - 1)) * chartW;
      const y = padT + chartH - (m.avgRating / maxR) * chartH;
      if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
    });
    ctx.stroke();

    meetings.forEach((m, i) => {
      const x = padL + (i / (meetings.length - 1)) * chartW;
      const y = padT + chartH - (m.avgRating / maxR) * chartH;
      ctx.beginPath();
      ctx.arc(x, y, 3, 0, Math.PI * 2);
      ctx.fillStyle = '#ff8c42';
      ctx.fill();
    });
  }

  render();
  new ResizeObserver(() => render()).observe(canvas.parentElement!);
}

function wireContentEvents(deptName: string): void {
  // Add person
  document.getElementById('btnAddPerson')?.addEventListener('click', () => {
    const list = document.getElementById('peopleList')!;
    const emptyMsg = list.querySelector('.admin-empty');
    if (emptyMsg) emptyMsg.remove();

    const div = document.createElement('div');
    div.className = 'people-item';
    const idx = list.querySelectorAll('.people-item').length;
    div.dataset.index = String(idx);
    div.innerHTML = `
      <input class="people-input" value="" placeholder="Name">
      <button class="people-remove" data-index="${idx}">&times;</button>
    `;
    list.appendChild(div);
    div.querySelector('input')?.focus();
    div.querySelector('.people-remove')?.addEventListener('click', () => {
      div.remove();
      savePeopleDebounced(deptName);
    });
  });

  // Remove person buttons
  document.querySelectorAll('.people-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      (btn as HTMLElement).closest('.people-item')?.remove();
      savePeopleDebounced(deptName);
    });
  });

  // Auto-save people on input changes
  const peopleList = document.getElementById('peopleList');
  if (peopleList) {
    peopleList.addEventListener('input', () => savePeopleDebounced(deptName));
  }

  // Import CSV
  document.getElementById('btnImportCsv')?.addEventListener('click', () => {
    (document.getElementById('csvFileInput') as HTMLInputElement)?.click();
  });
  document.getElementById('csvFileInput')?.addEventListener('change', async (e) => {
    const file = (e.target as HTMLInputElement).files?.[0];
    if (!file) return;
    try {
      const result = await fs.importMeetingFile(deptName, await file.arrayBuffer());
      if (!result) { alert('Import failed'); return; }
      location.hash = `#/dept/${encodeURIComponent(deptName)}/meeting/${result.id}`;
    } catch (err: any) {
      alert(`Import failed: ${err.message || err}`);
    }
    (e.target as HTMLInputElement).value = '';
  });

  // New meeting
  document.getElementById('btnNewMeeting')?.addEventListener('click', () => {
    location.hash = `#/dept/${encodeURIComponent(deptName)}/meeting/new`;
  });

  // Meeting item clicks
  document.querySelectorAll<HTMLElement>('.meeting-item').forEach(item => {
    item.addEventListener('click', (e) => {
      if ((e.target as HTMLElement).closest('.meeting-item-delete')) return;
      location.hash = `#/dept/${encodeURIComponent(deptName)}/meeting/${item.dataset.id}`;
    });
  });

  // Meeting delete buttons
  document.querySelectorAll<HTMLElement>('.meeting-item-delete').forEach(btn => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const id = (btn as HTMLElement).dataset.id!;
      if (!await confirmDialog('Delete this meeting? This cannot be undone.', 'Delete', true)) return;
      const ok = await fs.deleteMeeting(deptName, id);
      if (ok) {
        loadDepartmentContent(deptName);
        const sidebarItem = document.querySelector<HTMLElement>(`.sidebar-item[data-dept="${deptName}"] .sidebar-meta`);
        if (sidebarItem) {
          const count = parseInt(sidebarItem.textContent || '0');
          sidebarItem.textContent = String(Math.max(0, count - 1));
        }
      } else {
        alert('Failed to delete meeting');
      }
    });
  });

  // Rename department
  document.getElementById('btnRenameDept')?.addEventListener('click', async () => {
    const newName = prompt('New department name:', deptName);
    if (!newName || !newName.trim() || newName.trim() === deptName) return;
    const result = await fs.renameDepartment(deptName, newName.trim());
    if (!result.ok) {
      alert(result.error || 'Failed to rename');
      return;
    }
    await renderAdminPortal(newName.trim());
  });

  // Delete department
  document.getElementById('btnDeleteDept')?.addEventListener('click', async () => {
    if (!await confirmDialog(`Delete department "${deptName}" and all its meetings? This cannot be undone.`, 'Delete', true)) return;
    const ok = await fs.deleteDepartment(deptName);
    if (ok.ok) {
      _selectedDept = null;
      location.hash = '#/';
      await renderAdminPortal();
    } else {
      alert('Failed to delete department');
    }
  });
}

// Keep renderDepartmentView for the router -- it just calls renderAdminPortal with selection
export async function renderDepartmentView(deptName: string): Promise<void> {
  await renderAdminPortal(deptName);
}

// ── About Page ──

const FAQ_ITEMS: { q: string; a: string }[] = [
  {
    q: 'What is a Level 10 Meeting?',
    a: 'A Level 10 (L10) Meeting is a weekly leadership team meeting format from the Entrepreneurial Operating System (EOS). It follows a structured agenda designed to keep teams aligned, surface issues, and drive accountability. The name comes from the goal: every meeting should be rated a "10" by attendees.',
  },
  {
    q: 'How long does an L10 meeting take?',
    a: 'Exactly 90 minutes. The strict time-boxing is intentional — it forces prioritization and keeps discussions focused. Each section has a recommended time allocation that totals 90 minutes.',
  },
  {
    q: 'What are the 7 sections of an L10 meeting?',
    a: '<strong>1. Segue (5 min)</strong> — Share personal and professional good news.<br><strong>2. Scorecard Review (5 min)</strong> — Review weekly KPIs.<br><strong>3. OKR / Rock Review (5 min)</strong> — Report on/off track for quarterly goals.<br><strong>4. Headlines (5 min)</strong> — Customer and employee news.<br><strong>5. To-Do Review (5 min)</strong> — Check last week\'s commitments.<br><strong>6. IDS (60 min)</strong> — Identify, Discuss, Solve the top issues.<br><strong>7. Conclude (5 min)</strong> — Recap to-dos, cascading messages, and rate the meeting.',
  },
  {
    q: 'What does IDS stand for?',
    a: '<strong>Identify</strong> the real issue (not the symptom), <strong>Discuss</strong> it openly (ask "why?" until you find the root cause), and <strong>Solve</strong> it with a concrete action item. The key rule: solve one issue completely before moving to the next.',
  },
  {
    q: 'What is a Scorecard?',
    a: 'A scorecard tracks 5–15 weekly measurables (KPIs) that give your team a pulse on the business. Each measurable has an owner, a goal, and a weekly actual. If a number is off track, it drops into the Issues List for IDS.',
  },
  {
    q: 'What are Rocks / OKRs?',
    a: 'Rocks are the 3–7 most important priorities for the quarter (90 days). In this app they are tracked as OKRs (Objectives and Key Results). Each Rock has an owner and is reported as On Track or Off Track each week. Off-track Rocks go to IDS.',
  },
  {
    q: 'Who should attend the L10?',
    a: 'Your leadership team — typically 3–8 people. Everyone who has accountability for a departmental function should be in the room. Keep the group small enough for productive IDS conversations.',
  },
  {
    q: 'What is the meeting rating for?',
    a: 'At the end of each L10, every attendee rates the meeting from 1–10. The team discusses any rating below 8. Over time, this feedback loop improves meeting quality and highlights process issues.',
  },
  {
    q: 'What carries over from the previous meeting?',
    a: 'When you start a new meeting, the following data is automatically carried forward from the most recent meeting:<br><strong>Scorecard</strong> — KPI names, owners, and goals.<br><strong>OKR / Rock Review</strong> — Descriptions and owners.<br><strong>Scorecard Tracker</strong> — Full rolling 13-week data.<br><strong>OKR Tracker</strong> — Full OKR data and key results.<br><strong>Issues List</strong> — All issues from the IDS section.<br><strong>To-Do Review</strong> — Last week\'s new to-dos become this week\'s review items (with status carried over).',
  },
  {
    q: 'How does the Excel template work?',
    a: 'The app uses a blank Excel template as the canonical format. Every meeting is saved as a .xlsx file with a fixed layout. You can download the template from the sidebar and open it in Excel independently. When you import an Excel file, it must follow the same layout.',
  },
  {
    q: 'Where is my data stored?',
    a: 'All data stays on your local machine in the folder you selected when you first launched the app. Nothing is sent to any server. For backup, point the app at a folder synced with OneDrive, Google Drive, or Dropbox.',
  },
];

export function showAboutDialog(): void {
  // Don't open twice
  if (document.querySelector('.about-overlay')) return;

  const faqHtml = FAQ_ITEMS.map((item, i) => `
    <div class="about-faq-item">
      <button class="about-faq-q" data-faq="${i}">
        <span>${item.q}</span>
        <span class="about-faq-chevron">&#9662;</span>
      </button>
      <div class="about-faq-a" id="faq-a-${i}">${item.a}</div>
    </div>
  `).join('');

  const overlay = document.createElement('div');
  overlay.className = 'about-overlay';
  overlay.innerHTML = `
    <div class="about-dialog">
      <button class="about-dialog-close" title="Close">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </button>
      <div class="about-dialog-body">

        <div class="about-hero">
          <h1>About L10 Meetings</h1>
          <p>The Level 10 Meeting is the heartbeat of a team running on the Entrepreneurial Operating System (EOS). It turns meetings from time-wasters into the most productive 90 minutes of your week. <a href="https://www.eosworldwide.com/blog/three-steps-improve-level-ten-meeting" target="_blank" rel="noopener" class="about-learn-more">Learn more</a></p>
        </div>

        <div class="about-section">
          <h2>How It Works</h2>
          <div class="about-steps">
            <div class="about-step">
              <div class="about-step-num">1</div>
              <div>
                <strong>Same day, same time, every week</strong>
                <p>Consistency builds rhythm. The meeting starts and ends on time — no exceptions.</p>
              </div>
            </div>
            <div class="about-step">
              <div class="about-step-num">2</div>
              <div>
                <strong>Follow the agenda</strong>
                <p>Seven sections, time-boxed. Segue, Scorecard, OKR Review, Headlines, To-Do Review, IDS, and Conclude.</p>
              </div>
            </div>
            <div class="about-step">
              <div class="about-step-num">3</div>
              <div>
                <strong>IDS is the engine</strong>
                <p>60 of 90 minutes go to Identify, Discuss, Solve. Surface the real issue, find the root cause, and agree on a concrete next step.</p>
              </div>
            </div>
            <div class="about-step">
              <div class="about-step-num">4</div>
              <div>
                <strong>End with accountability</strong>
                <p>Recap new to-dos, decide what to cascade to the organization, and rate the meeting 1–10.</p>
              </div>
            </div>
          </div>
        </div>

        <div class="about-section">
          <h2>Frequently Asked Questions</h2>
          ${faqHtml}
        </div>

      </div>
    </div>
  `;

  document.body.appendChild(overlay);
  requestAnimationFrame(() => overlay.classList.add('visible'));

  function close() {
    overlay.classList.remove('visible');
    overlay.addEventListener('transitionend', () => overlay.remove(), { once: true });
  }

  // Close button
  overlay.querySelector('.about-dialog-close')!.addEventListener('click', close);

  // Click outside dialog to close
  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) close();
  });

  // Escape key
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') { close(); document.removeEventListener('keydown', onKey); }
  };
  document.addEventListener('keydown', onKey);

  // FAQ accordion
  overlay.querySelectorAll<HTMLButtonElement>('.about-faq-q').forEach(btn => {
    btn.addEventListener('click', () => {
      const i = btn.dataset.faq!;
      const answer = document.getElementById(`faq-a-${i}`)!;
      const isOpen = answer.classList.contains('open');
      overlay.querySelectorAll('.about-faq-a').forEach(a => a.classList.remove('open'));
      overlay.querySelectorAll('.about-faq-chevron').forEach(c => c.classList.remove('open'));
      if (!isOpen) {
        answer.classList.add('open');
        btn.querySelector('.about-faq-chevron')!.classList.add('open');
      }
    });
  });
}
