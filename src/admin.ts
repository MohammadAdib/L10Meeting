import logoUrl from './logo.png';

export async function renderAdminPortal(): Promise<void> {
  const app = document.getElementById('app')!;

  let departments: { name: string; meetingCount: number }[] = [];
  try {
    const res = await fetch('/api/departments');
    departments = await res.json();
  } catch { /* empty */ }

  const cards = departments.map(d => `
    <div class="dept-card" data-dept="${d.name}">
      <div class="dept-card-name">${d.name}</div>
      <div class="dept-card-meta">${d.meetingCount} meeting${d.meetingCount !== 1 ? 's' : ''}</div>
    </div>
  `).join('');

  app.innerHTML = `
    <div class="top-bar-wrapper">
      <div class="top-bar">
        <div class="top-bar-left">
          <img src="${logoUrl}" alt="Titan Dynamics" class="top-bar-logo">
          <div class="top-bar-title">L10 Meeting Manager</div>
        </div>
        <div class="top-bar-actions" style="opacity:1;pointer-events:auto">
        </div>
      </div>
    </div>
    <div class="admin-container">
      <div class="admin-header">
        <h1>Departments</h1>
        <button class="btn btn-primary" id="btnAddDept">+ Add Department</button>
      </div>
      <div class="admin-grid">
        ${cards || '<div class="admin-empty">No departments yet. Create one to get started.</div>'}
      </div>
    </div>
    <div class="toast"></div>
  `;

  // Wire up card clicks
  app.querySelectorAll<HTMLElement>('.dept-card').forEach(card => {
    card.addEventListener('click', () => {
      location.hash = `#/dept/${encodeURIComponent(card.dataset.dept!)}`;
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
      location.hash = `#/dept/${encodeURIComponent(name.trim())}`;
    } catch {
      alert('Failed to create department');
    }
  });
}

export async function renderDepartmentView(deptName: string): Promise<void> {
  const app = document.getElementById('app')!;

  // Fetch people and meetings in parallel
  let people: string[] = [];
  let meetings: { id: string; date: string; lastSaved: string }[] = [];

  try {
    const [pRes, mRes] = await Promise.all([
      fetch(`/api/departments/${encodeURIComponent(deptName)}/people`),
      fetch(`/api/departments/${encodeURIComponent(deptName)}/meetings`),
    ]);
    people = await pRes.json();
    meetings = await mRes.json();
  } catch { /* empty */ }

  const peopleItems = people.map((p, i) => `
    <div class="people-item" data-index="${i}">
      <input class="people-input" value="${p.replace(/"/g, '&quot;')}" placeholder="Name">
      <button class="people-remove" data-index="${i}">&times;</button>
    </div>
  `).join('');

  const meetingItems = meetings.map(m => {
    const lastSaved = m.lastSaved ? new Date(m.lastSaved).toLocaleString() : '';
    return `
      <div class="meeting-item" data-id="${m.id}">
        <div class="meeting-item-date">${m.date}</div>
        <div class="meeting-item-saved">${lastSaved ? `Last saved: ${lastSaved}` : ''}</div>
        <button class="meeting-item-delete btn btn-outline-dark btn-sm" data-id="${m.id}" title="Delete meeting">&times;</button>
      </div>
    `;
  }).join('');

  app.innerHTML = `
    <div class="top-bar-wrapper">
      <div class="top-bar">
        <div class="top-bar-left">
          <button class="back-btn" id="btnBack" title="Back to departments">&#8592;</button>
          <img src="${logoUrl}" alt="Titan Dynamics" class="top-bar-logo">
          <div class="top-bar-title">${deptName}</div>
        </div>
        <div class="top-bar-actions" style="opacity:1;pointer-events:auto">
          <button class="btn btn-outline" id="btnRenameDept">Rename</button>
          <button class="btn btn-outline" id="btnDeleteDept" style="border-color:var(--red);color:var(--red);">Delete Dept</button>
        </div>
      </div>
    </div>
    <div class="admin-container">
      <!-- People Section -->
      <div class="dept-section">
        <div class="dept-section-header">
          <h2>People</h2>
          <button class="btn btn-outline-dark btn-sm" id="btnAddPerson">+ Add Person</button>
        </div>
        <div class="people-list" id="peopleList">
          ${peopleItems || '<div class="admin-empty" style="padding:12px;">No people added yet.</div>'}
        </div>
        <button class="btn btn-primary btn-sm" id="btnSavePeople" style="margin-top:8px;">Save People</button>
      </div>

      <!-- Meetings Section -->
      <div class="dept-section">
        <div class="dept-section-header">
          <h2>Meetings</h2>
          <button class="btn btn-primary" id="btnNewMeeting">+ New L10 Meeting</button>
        </div>
        <div class="meetings-list">
          ${meetingItems || '<div class="admin-empty" style="padding:12px;">No meetings yet.</div>'}
        </div>
      </div>
    </div>
    <div class="toast"></div>
  `;

  // Wire events
  document.getElementById('btnBack')?.addEventListener('click', () => {
    location.hash = '#/';
  });

  // Add person
  document.getElementById('btnAddPerson')?.addEventListener('click', () => {
    const list = document.getElementById('peopleList')!;
    // Remove empty message if present
    const emptyMsg = list.querySelector('.admin-empty');
    if (emptyMsg) emptyMsg.remove();

    const div = document.createElement('div');
    div.className = 'people-item';
    const idx = list.querySelectorAll('.people-item').length;
    div.dataset.index = String(idx);
    div.innerHTML = `
      <input class="people-input" value="" placeholder="Name" autofocus>
      <button class="people-remove" data-index="${idx}">&times;</button>
    `;
    list.appendChild(div);
    div.querySelector('input')?.focus();

    div.querySelector('.people-remove')?.addEventListener('click', () => {
      div.remove();
    });
  });

  // Remove person buttons
  document.querySelectorAll('.people-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      (btn as HTMLElement).closest('.people-item')?.remove();
    });
  });

  // Save people
  document.getElementById('btnSavePeople')?.addEventListener('click', async () => {
    const inputs = document.querySelectorAll<HTMLInputElement>('#peopleList .people-input');
    const names = Array.from(inputs).map(i => i.value.trim()).filter(Boolean);
    try {
      await fetch(`/api/departments/${encodeURIComponent(deptName)}/people`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ people: names }),
      });
      showToastSimple('People saved!');
    } catch {
      showToastSimple('Error saving people');
    }
  });

  // New meeting
  document.getElementById('btnNewMeeting')?.addEventListener('click', async () => {
    try {
      const res = await fetch(`/api/departments/${encodeURIComponent(deptName)}/meetings`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({}),
      });
      const { id } = await res.json();
      location.hash = `#/dept/${encodeURIComponent(deptName)}/meeting/${id}`;
    } catch {
      alert('Failed to create meeting');
    }
  });

  // Meeting item clicks
  document.querySelectorAll<HTMLElement>('.meeting-item').forEach(item => {
    item.addEventListener('click', (e) => {
      // Don't navigate if clicking delete button
      if ((e.target as HTMLElement).closest('.meeting-item-delete')) return;
      location.hash = `#/dept/${encodeURIComponent(deptName)}/meeting/${item.dataset.id}`;
    });
  });

  // Meeting delete buttons
  document.querySelectorAll<HTMLElement>('.meeting-item-delete').forEach(btn => {
    btn.addEventListener('click', async (e) => {
      e.stopPropagation();
      const id = (btn as HTMLElement).dataset.id!;
      if (!confirm(`Delete meeting ${id}?`)) return;
      try {
        await fetch(`/api/departments/${encodeURIComponent(deptName)}/meetings/${id}`, {
          method: 'DELETE',
        });
        renderDepartmentView(deptName); // refresh
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
      location.hash = `#/dept/${encodeURIComponent(newName.trim())}`;
    } catch {
      alert('Failed to rename department');
    }
  });

  // Delete department
  document.getElementById('btnDeleteDept')?.addEventListener('click', async () => {
    if (!confirm(`Delete department "${deptName}" and ALL its meetings? This cannot be undone.`)) return;
    try {
      await fetch(`/api/departments/${encodeURIComponent(deptName)}`, { method: 'DELETE' });
      location.hash = '#/';
    } catch {
      alert('Failed to delete department');
    }
  });
}

function showToastSimple(msg: string): void {
  const t = document.querySelector<HTMLElement>('.toast');
  if (!t) return;
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2500);
}
