import logoUrl from './logo.png';
import { confirmDialog } from './utils';
import { buildScorecardContent, buildOkrsContent } from './html';
import { addScorecardFullRow, addOkrFullRow, buildKeyResultBlocks } from './tables';
import { loadScorecardOkrData } from './storage';

let _selectedDept: string | null = null;
let _peopleSaveTimer: ReturnType<typeof setTimeout> | null = null;

function savePeopleDebounced(deptName: string): void {
  if (_peopleSaveTimer) clearTimeout(_peopleSaveTimer);
  _peopleSaveTimer = setTimeout(async () => {
    const inputs = document.querySelectorAll<HTMLInputElement>('#peopleList .people-input');
    const names = Array.from(inputs).map(i => i.value.trim()).filter(Boolean);
    try {
      await fetch(`/api/departments/${encodeURIComponent(deptName)}/people`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ people: names }),
      });
    } catch { /* silent */ }
  }, 1000);
}

export async function renderAdminPortal(selectedDept?: string): Promise<void> {
  _selectedDept = selectedDept || null;
  const app = document.getElementById('app')!;

  let departments: { name: string; peopleCount: number }[] = [];
  try {
    const res = await fetch('/api/departments');
    departments = await res.json();
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
          <img src="${logoUrl}" alt="Titan Dynamics" class="top-bar-logo">
          <div class="top-bar-title">L10 Meeting Manager</div>
        </div>
        <div class="top-bar-actions" style="opacity:1;pointer-events:auto"></div>
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
      location.hash = `#/dept/${encodeURIComponent(dept)}`;
      loadDepartmentContent(dept);
    });
  });

  // Add department button
  document.getElementById('btnAddDept')?.addEventListener('click', async () => {
    const name = prompt('Department name:');
    if (!name || !name.trim()) return;
    try {
      const res = await fetch('/api/departments', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: name.trim() }),
      });
      if (!res.ok) {
        const err = await res.json();
        alert(err.error || 'Failed to create department');
        return;
      }
      // Re-render with new dept selected
      await renderAdminPortal(name.trim());
    } catch {
      alert('Failed to create department');
    }
  });

  // Load selected department content
  if (_selectedDept) {
    await loadDepartmentContent(_selectedDept);
  }
}

async function loadDepartmentContent(deptName: string): Promise<void> {
  const content = document.getElementById('adminContent');
  if (!content) return;

  // Trigger fade-slide animation on content swap
  content.style.animation = 'none';
  content.offsetHeight; // force reflow
  content.style.animation = 'fadeSlideRight .3s ease forwards';

  let people: string[] = [];
  let meetings: { id: string; date: string; lastSaved: string; avgRating: number }[] = [];

  try {
    const [pRes, mRes] = await Promise.all([
      fetch(`/api/departments/${encodeURIComponent(deptName)}/people`),
      fetch(`/api/departments/${encodeURIComponent(deptName)}/meetings`),
    ]);
    people = await pRes.json();
    meetings = await mRes.json();
  } catch { /* empty */ }

  // Load last meeting data for scorecard/OKR (read-only)
  let lastMeetingData: Record<string, unknown> | null = null;
  if (meetings.length > 0) {
    try {
      const res = await fetch(`/api/departments/${encodeURIComponent(deptName)}/meetings/${meetings[0].id}`);
      if (res.ok) lastMeetingData = await res.json();
    } catch { /* empty */ }
  }

  const scorecardRows = (lastMeetingData?.scorecardFullTable as string[][] | undefined) || [];
  const okrRows = (lastMeetingData?.okrFullTable as string[][] | undefined) || [];

  const peopleItems = people.map((p, i) => `
    <div class="people-item" data-index="${i}">
      <input class="people-input" value="${p.replace(/"/g, '&quot;')}" placeholder="Name">
      <button class="people-remove" data-index="${i}">&times;</button>
    </div>
  `).join('');

  const meetingItems = meetings.map(m => {
    // Format date as dd/mm/yyyy
    let displayDate = m.date;
    const dp = m.date.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (dp) displayDate = `${parseInt(dp[2])}/${dp[3]}/${dp[1]}`;

    // Star rating display
    const avg = m.avgRating || 0;
    const fullStars = Math.round(avg);
    const starsHtml = avg > 0
      ? Array.from({ length: 10 }, (_, i) =>
          `<span class="meeting-star${i < fullStars ? ' active' : ''}">\u2605</span>`
        ).join('') + `<span class="meeting-rating-val">${avg.toFixed(1)}</span>`
      : '<span class="meeting-rating-val" style="color:var(--text-muted)">No ratings</span>';

    return `
      <div class="meeting-item" data-id="${m.id}">
        <div class="meeting-item-date">${displayDate}</div>
        <div class="meeting-item-rating">${starsHtml}</div>
        <button class="meeting-item-delete btn btn-outline-dark btn-sm" data-id="${m.id}" title="Delete meeting">&times;</button>
      </div>
    `;
  }).join('');

  content.innerHTML = `
    <div class="dept-header">
      <h1>${deptName}</h1>
      <div class="dept-header-actions">
        <button class="btn btn-outline btn-sm" id="btnRenameDept">Rename</button>
        <button class="btn btn-outline btn-sm" id="btnDeleteDept" style="border-color:var(--red);color:var(--red);">Delete</button>
      </div>
    </div>

    <div class="dept-layout">
      <div class="dept-left">
        <div class="dept-section">
          <div class="dept-section-header">
            <h2>People</h2>
            <button class="btn btn-outline-dark btn-sm" id="btnAddPerson">+ Add Person</button>
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
        </div>
      </div>

      <div class="dept-right dept-readonly${meetings.length === 0 ? ' dept-no-meetings' : ''}">
        <div class="dept-readonly-label">${meetings.length > 0 ? `Scorecard and OKRs from most recent meeting (${meetings[0].date})` : 'Scorecard and OKRs from most recent meeting — no meetings yet'}</div>
        ${buildScorecardContent()}
        ${buildOkrsContent()}
      </div>
    </div>
  `;

  // Populate scorecard/OKR rows — use defaults or previous meeting count
  const scCount = scorecardRows.length || 3;
  for (let i = 0; i < scCount; i++) addScorecardFullRow();

  const okCount = okrRows.length || 3;
  for (let i = 1; i <= okCount; i++) addOkrFullRow('', i);
  buildKeyResultBlocks();

  // Load data from last meeting using shared loader
  if (lastMeetingData) {
    loadScorecardOkrData(lastMeetingData as Record<string, unknown>);
  }

  // Make dept-right read-only: disable all inputs/selects and hide add/delete buttons
  const deptRight = document.querySelector('.dept-right');
  if (deptRight) {
    deptRight.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea').forEach(el => {
      el.disabled = true;
    });
    deptRight.querySelectorAll<HTMLElement>('.add-row-btn, .row-delete').forEach(el => {
      el.style.display = 'none';
    });
  }

  wireContentEvents(deptName);
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
      const res = await fetch(`/api/departments/${encodeURIComponent(deptName)}/meetings/import`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/octet-stream' },
        body: await file.arrayBuffer(),
      });
      const result = await res.json();
      if (!res.ok) { alert(`Import failed: ${result.error || 'Unknown error'}`); return; }
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
      try {
        await fetch(`/api/departments/${encodeURIComponent(deptName)}/meetings/${id}`, { method: 'DELETE' });
        loadDepartmentContent(deptName);
        // Refresh sidebar count
        const sidebarItem = document.querySelector<HTMLElement>(`.sidebar-item[data-dept="${deptName}"] .sidebar-meta`);
        if (sidebarItem) {
          const count = parseInt(sidebarItem.textContent || '0');
          sidebarItem.textContent = String(Math.max(0, count - 1));
        }
      } catch {
        alert('Failed to delete meeting');
      }
    });
  });

  // Rename department
  document.getElementById('btnRenameDept')?.addEventListener('click', async () => {
    const newName = prompt('New department name:', deptName);
    if (!newName || !newName.trim() || newName.trim() === deptName) return;
    try {
      const res = await fetch(`/api/departments/${encodeURIComponent(deptName)}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: newName.trim() }),
      });
      if (!res.ok) {
        const err = await res.json();
        alert(err.error || 'Failed to rename');
        return;
      }
      await renderAdminPortal(newName.trim());
    } catch {
      alert('Failed to rename department');
    }
  });

  // Delete department
  document.getElementById('btnDeleteDept')?.addEventListener('click', async () => {
    if (!await confirmDialog(`Delete department "${deptName}" and all its meetings? This cannot be undone.`, 'Delete', true)) return;
    try {
      await fetch(`/api/departments/${encodeURIComponent(deptName)}`, { method: 'DELETE' });
      _selectedDept = null;
      location.hash = '#/';
      await renderAdminPortal();
    } catch {
      alert('Failed to delete department');
    }
  });
}

// Keep renderDepartmentView for the router -- it just calls renderAdminPortal with selection
export async function renderDepartmentView(deptName: string): Promise<void> {
  await renderAdminPortal(deptName);
}
