import { getLogoUrl, handleLogoClick } from './logo';
import { confirmDialog } from './utils';
import { buildScorecardContent, buildOkrsContent } from './html';
import { addScorecardFullRow, addOkrFullRow, buildKeyResultBlocks } from './tables';
import { loadScorecardOkrData } from './storage';
import * as fs from './fs-service';
import { showSettingsMenu } from './settings';
import { DEFAULT_ROWS } from './types';

let _selectedDept: string | null = null;
let _peopleSaveTimer: ReturnType<typeof setTimeout> | null = null;

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
  if (avg <= 0) return '<span class="meeting-rating-val" style="color:var(--text-muted)">—</span>';
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
        <div class="meeting-item-rating" id="rating-${m.id}">${buildRatingHtml(m.avgRating)}</div>
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
  let meetings: { id: string; date: string; lastSaved: string; avgRating: number }[] = [];

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

      <div class="dept-right dept-readonly${meetings.length === 0 ? ' dept-no-meetings' : ' dept-loading'}">
        <div class="dept-readonly-label">${meetings.length > 0 ? `Scorecard and OKRs from most recent meeting (${meetings[0].date})` : 'Scorecard and OKRs from most recent meeting — no meetings yet'}</div>
        ${buildScorecardContent()}
        ${buildOkrsContent()}
      </div>
    </div>
  `;

  wireContentEvents(deptName);

  // ── Populate scorecard/OKR ──
  if (meetings.length === 0) {
    // No meetings — show default empty rows (read-only)
    for (let i = 0; i < DEFAULT_ROWS.scorecard; i++) addScorecardFullRow();
    for (let i = 1; i <= DEFAULT_ROWS.okr; i++) addOkrFullRow('', i);
    buildKeyResultBlocks();
    const dr = document.querySelector('.dept-right');
    if (dr) {
      dr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea').forEach(el => {
        el.disabled = true;
      });
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

      const scCount = Math.max(scorecardRows.length, DEFAULT_ROWS.scorecard);
      for (let i = 0; i < scCount; i++) addScorecardFullRow();
      const okCount = Math.max(okrRows.length, DEFAULT_ROWS.okr);
      for (let i = 1; i <= okCount; i++) addOkrFullRow('', i);
      buildKeyResultBlocks();

      if (lastMeetingData) {
        loadScorecardOkrData(lastMeetingData as Record<string, unknown>);
      }

      // Apply read-only and reveal
      const dr = document.querySelector<HTMLElement>('.dept-right');
      if (dr) {
        dr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea').forEach(el => {
          el.disabled = true;
        });
        dr.querySelectorAll<HTMLElement>('.add-row-btn, .row-delete').forEach(el => {
          el.style.display = 'none';
        });
        dr.classList.remove('dept-loading');
      }
    });

    // ── Background: load meeting ratings ──
    fs.loadMeetingRatings(deptName).then(rated => {
      if (_selectedDept !== deptName) return;
      for (const m of rated) {
        const el = document.getElementById(`rating-${m.id}`);
        if (el) el.innerHTML = buildRatingHtml(m.avgRating);
      }
    });
  }
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
