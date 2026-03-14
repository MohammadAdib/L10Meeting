import './style.css';
import { buildAppHTML } from './html';
import { initTimers, toggleTimer, resetTimer, cleanupTimers } from './timer';
import { onStatusChange, confirmDialog, initPersonPickers } from './utils';
import { loadMeetingData, loadScorecardOkrData, setupAutoSave, markMeetingStarted, markMeetingStopped, isMeetingActive, cleanupAutoSave, disableAutoSave, forceSave, setupScorecardOkrSync } from './storage';
import { DEFAULT_MEASURABLES, DEFAULT_ROWS } from './types';
import { renderAdminPortal, renderDepartmentView } from './admin';
import {
  addScorecardRow, addOkrReviewRow, addHeadlineRow, addTodoReviewRow,
  addIssueRow, addIDSIssue, addIDSTodoRow, addNewTodoRow, addCascadingRow,
  addRatingRow, setRating, updateTodoCompletion, updateAvgRating,
  addScorecardFullRow, addOkrFullRow, addKeyResultRow, buildKeyResultBlocks,
  resetIdsIssueCount, setPeople, showCapToast,
} from './tables';
import * as fs from './fs-service';
import { getLogoUrl, initLogo, handleLogoClick } from './logo';

// ── Expose globals for inline onclick handlers ──
declare global {
  interface Window {
    __onStatusChange: typeof onStatusChange;
    __updateTodoCompletion: typeof updateTodoCompletion;
    __updateAvgRating: typeof updateAvgRating;
    __setRating: typeof setRating;
    __addIDSTodoRow: typeof addIDSTodoRow;
    __addKeyResultRow: typeof addKeyResultRow;
  }
}
window.__onStatusChange = onStatusChange;
window.__updateTodoCompletion = updateTodoCompletion;
window.__updateAvgRating = updateAvgRating;
window.__setRating = setRating;
window.__addIDSTodoRow = (n: number) => {
  const before = document.querySelectorAll(`#idsTodo-${n} tbody tr`).length;
  addIDSTodoRow(n);
  if (document.querySelectorAll(`#idsTodo-${n} tbody tr`).length === before) showCapToast();
};
window.__addKeyResultRow = (n: number) => {
  const before = document.querySelectorAll(`#keyResults-${n} tbody tr`).length;
  addKeyResultRow(n);
  if (document.querySelectorAll(`#keyResults-${n} tbody tr`).length === before) showCapToast();
};

// ── Router ──
let _previousHash = '';
async function route() {
  const hash = location.hash || '#/';
  const leavingMeeting = _previousHash.includes('/meeting/') && !hash.includes('/meeting/');

  // Confirm before leaving an active meeting
  if (leavingMeeting && isMeetingActive()) {
    if (!await confirmDialog('You have an active meeting. Are you sure you want to leave?', 'Leave')) {
      // Restore the previous hash without triggering another route
      history.pushState(null, '', _previousHash);
      return;
    }
  }

  _previousHash = hash;

  // Clean up from any previous meeting view
  cleanupTimers();
  await cleanupAutoSave();

  const meetingMatch = hash.match(/#\/dept\/([^/]+)\/meeting\/(.+)/);
  const deptMatch = hash.match(/#\/dept\/([^/]+)$/);

  if (meetingMatch) {
    const deptName = decodeURIComponent(meetingMatch[1]);
    const meetingId = decodeURIComponent(meetingMatch[2]);
    await initMeetingView(deptName, meetingId);
  } else if (deptMatch) {
    const deptName = decodeURIComponent(deptMatch[1]);
    await renderDepartmentView(deptName);
  } else {
    await renderAdminPortal();
  }
}

async function initMeetingView(deptName: string, meetingId: string): Promise<void> {
  const app = document.getElementById('app')!;

  // Reset IDS issue counter for fresh meeting
  resetIdsIssueCount();

  // Render the meeting HTML
  app.innerHTML = buildAppHTML(deptName);

  // ── Auto-fill date ──
  const now = new Date();
  (document.getElementById('metaDate') as HTMLInputElement).value = now.toISOString().split('T')[0];

  // ── Pre-fill team name with department ──
  (document.getElementById('metaTeam') as HTMLInputElement).value = deptName;

  // ── Duration helper ──
  function updateDuration(): void {
    const controlDiv = document.querySelector('.meeting-control');
    if (!controlDiv) return;
    const startVal = (document.getElementById('metaStart') as HTMLInputElement)?.value;
    const endVal = (document.getElementById('metaEnd') as HTMLInputElement)?.value;
    if (!startVal || !endVal) {
      controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">&nbsp;</span>`;
      return;
    }
    const sp = startVal.split(':').map(Number);
    const ep = endVal.split(':').map(Number);
    if (sp.length < 2 || ep.length < 2 || sp.some(isNaN) || ep.some(isNaN)) {
      controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">&nbsp;</span>`;
      return;
    }
    const totalMins = (ep[0] * 60 + ep[1]) - (sp[0] * 60 + sp[1]);
    if (totalMins <= 0) {
      controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">&nbsp;</span>`;
      return;
    }
    const h = Math.floor(totalMins / 60);
    const m = totalMins % 60;
    const elapsed = h > 0 ? `${h}:${String(m).padStart(2, '0')}:00` : `${m}:00`;
    controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">Duration: ${elapsed}</span>`;
  }

  // ── Meeting start/stop ──
  const meetingTab = document.getElementById('tab-meeting')!;
  const sidebar = document.getElementById('sidebar')!;
  const isExisting = meetingId !== 'new';

  if (isExisting) {
    // Existing meeting: no blur, no start/stop, no section timers, show actions immediately
    const actions = document.getElementById('topBarActions')!;
    actions.style.opacity = '1';
    actions.style.pointerEvents = '';
    const controlDiv = document.querySelector('.meeting-control')!;
    controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">&nbsp;</span>`;
    document.querySelectorAll<HTMLElement>('.section-timer').forEach(el => el.style.display = 'none');
    document.querySelectorAll<HTMLElement>('.section-duration').forEach(el => el.style.display = 'inline');
    const btnDelete = document.getElementById('btnDeleteMeeting');
    if (btnDelete) {
      btnDelete.style.display = '';
      btnDelete.addEventListener('click', async () => {
        if (!await confirmDialog('Delete this meeting? This cannot be undone.', 'Delete', true)) return;
        const ok = await fs.deleteMeeting(deptName, meetingId);
        if (ok) {
          disableAutoSave();
          location.hash = `#/dept/${encodeURIComponent(deptName)}`;
        }
      });
    }
  } else {
    // New meeting: blur until started
    meetingTab.classList.add('blurred');
    sidebar.classList.add('blurred');

    let meetingInterval: ReturnType<typeof setInterval> | null = null;
    let meetingSeconds = 0;

    function formatElapsed(secs: number): string {
      const h = Math.floor(secs / 3600);
      const m = Math.floor((secs % 3600) / 60);
      const s = secs % 60;
      if (h > 0) return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
      return `${m}:${String(s).padStart(2, '0')}`;
    }

    function startMeeting() {
      markMeetingStarted();
      meetingTab.classList.remove('blurred');
      sidebar.classList.remove('blurred');
      const actions = document.getElementById('topBarActions')!;
      actions.style.opacity = '1';
      actions.style.pointerEvents = '';
      const startNow = new Date();
      (document.getElementById('metaStart') as HTMLInputElement).value = startNow.toTimeString().slice(0, 5);

      const controlDiv = document.querySelector('.meeting-control')!;
      meetingSeconds = 0;
      controlDiv.innerHTML = `
        <button class="meeting-stop-btn" id="btnMeetingStop">
          <span class="stop-icon"></span>
          <span class="meeting-timer-display" id="meetingElapsed">0:00</span>
        </button>`;

      meetingInterval = setInterval(() => {
        meetingSeconds++;
        const display = document.getElementById('meetingElapsed');
        if (display) display.textContent = formatElapsed(meetingSeconds);
      }, 1000);

      document.getElementById('btnMeetingStop')!.addEventListener('click', stopMeeting);
    }

    function stopMeeting() {
      if (meetingInterval) clearInterval(meetingInterval);
      markMeetingStopped();
      const endNow = new Date();
      (document.getElementById('metaEnd') as HTMLInputElement).value = endNow.toTimeString().slice(0, 5);
      updateDuration();
      forceSave();
    }

    document.getElementById('btnMeetingStart')!.addEventListener('click', startMeeting);
  }

  // ── Init timers ──
  initTimers();

  // ── Fetch department people for dropdowns ──
  let people: string[] = [];
  try {
    people = await fs.getPeople(deptName);
  } catch { /* empty */ }
  setPeople(people);

  // ── Convert meta Facilitator/Scribe to people dropdowns ──
  for (const id of ['metaFacilitator', 'metaScribe']) {
    const input = document.getElementById(id) as HTMLInputElement | null;
    if (!input) continue;
    const sel = document.createElement('select');
    sel.id = id;
    sel.className = 'person-select';
    sel.innerHTML = `<option value=""></option>` +
      people.map(p => `<option value="${p}">${p}</option>`).join('');
    input.replaceWith(sel);
  }

  // ── Populate default rows ──
  DEFAULT_MEASURABLES.forEach(m => addScorecardRow(m));
  for (let i = 0; i < DEFAULT_ROWS.okr; i++) addOkrReviewRow();
  for (let i = 0; i < DEFAULT_ROWS.headlines; i++) addHeadlineRow();
  for (let i = 0; i < DEFAULT_ROWS.todoReview; i++) addTodoReviewRow();
  for (let i = 0; i < DEFAULT_ROWS.issues; i++) addIssueRow();
  for (let i = 0; i < DEFAULT_ROWS.idsIssues; i++) addIDSIssue();
  for (let i = 0; i < DEFAULT_ROWS.newTodos; i++) addNewTodoRow();
  for (let i = 0; i < DEFAULT_ROWS.cascading; i++) addCascadingRow();

  // ── Pre-fill rating table with department people ──
  if (people.length > 0) {
    people.forEach(() => addRatingRow());
    // Pre-select each person in the rating name dropdown
    const selects = document.querySelectorAll<HTMLSelectElement>('#ratingTable tbody tr .person-select');
    people.forEach((name, i) => {
      if (i < selects.length) selects[i].value = name;
    });
  } else {
    for (let i = 0; i < DEFAULT_ROWS.rating; i++) addRatingRow();
  }

  DEFAULT_MEASURABLES.forEach(m => addScorecardFullRow(m));
  for (let i = 1; i <= DEFAULT_ROWS.okr; i++) addOkrFullRow('', i);
  buildKeyResultBlocks();

  // ── For new meetings, carry over scorecard & OKR data from most recent meeting ──
  if (meetingId === 'new') {
    try {
      const meetings = await fs.getMeetings(deptName);
      if (meetings.length > 0) {
        meetings.sort((a: any, b: any) => (b.lastSaved || '').localeCompare(a.lastSaved || ''));
        const lastId = meetings[0].id;
        const lastData = await fs.getMeetingData(deptName, lastId);
        if (lastData) {
          const scRows = (lastData.scorecardTable as string[][] | undefined)?.filter((r: string[]) => r.some(c => c));
          if (scRows && scRows.length > 0) {
            const scTbody = document.querySelector('#scorecardTable tbody');
            if (scTbody) scTbody.innerHTML = '';
            scRows.forEach((cells: string[]) => {
              addScorecardRow(cells[0] || '');
              const tr = document.querySelector('#scorecardTable tbody tr:last-child');
              if (!tr) return;
              const els = tr.querySelectorAll<HTMLInputElement | HTMLSelectElement>('input, select');
              [1, 2].forEach(ci => {
                if (ci < els.length && ci < cells.length) els[ci].value = cells[ci];
              });
            });
          }
          const okrRows = (lastData.okrReviewTable as string[][] | undefined)?.filter((r: string[]) => r.some(c => c));
          if (okrRows && okrRows.length > 0) {
            const okrTbody = document.querySelector('#okrReviewTable tbody');
            if (okrTbody) okrTbody.innerHTML = '';
            okrRows.forEach((cells: string[]) => {
              addOkrReviewRow(cells[0] || '');
              const tr = document.querySelector('#okrReviewTable tbody tr:last-child');
              if (!tr) return;
              const els = tr.querySelectorAll<HTMLInputElement | HTMLSelectElement>('input, select');
              [1, 2].forEach(ci => {
                if (ci < els.length && ci < cells.length) els[ci].value = cells[ci];
              });
            });
          }
          loadScorecardOkrData(lastData);
        }
      }
    } catch { /* silent */ }
  }

  // ── If loading existing meeting, populate data ──
  if (meetingId !== 'new') {
    let data: Record<string, any> | null = null;
    let parseError = false;
    try {
      data = await fs.getMeetingData(deptName, meetingId);
    } catch {
      parseError = true;
    }

    if (parseError || (data !== null && Object.keys(data).length === 0)) {
      // File exists but couldn't be parsed
      const container = document.querySelector('.main-content .container')!;
      container.innerHTML = `
        <div class="parse-error">
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
          </svg>
          <h2>Unable to render this meeting</h2>
          <p>The Excel file could not be parsed. It may have been modified in a way that doesn't match the expected L10 template layout.</p>
          <a class="btn btn-outline" href="#/dept/${encodeURIComponent(deptName)}" style="text-decoration:none;margin-top:8px">&larr; Back to department</a>
        </div>`;
      return;
    }

    if (data) {
      loadMeetingData(data);
      updateDuration();
    } else {
      location.replace('#/');
      return;
    }
  }

  // ── Set up auto-save ──
  // For new meetings, defer creating on server until first save
  setupAutoSave(deptName, meetingId === 'new' ? '' : meetingId, meetingId === 'new');
  setupScorecardOkrSync();

  // ── Recalc duration on time changes ──
  document.getElementById('metaStart')?.addEventListener('change', updateDuration);
  document.getElementById('metaEnd')?.addEventListener('change', updateDuration);

  // ── Event Delegation ──

  // Logo click → back to department view
  document.querySelector('.top-bar-logo')?.addEventListener('click', () => {
    location.hash = `#/dept/${encodeURIComponent(deptName)}`;
  });

  // Add logo button (if no logo set)
  document.getElementById('btnAddLogo')?.addEventListener('click', () => {
    handleLogoClick(() => initMeetingView(deptName, meetingId));
  });

  // Tab switching
  document.querySelectorAll<HTMLButtonElement>('.top-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab!;
      document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.top-tab').forEach(b => b.classList.remove('active'));
      document.getElementById(`tab-${tab}`)?.classList.add('active');
      btn.classList.add('active');
      // Show sidebar only on meeting tab
      const mainContent = document.querySelector<HTMLElement>('.main-content');
      if (mainContent) {
        sidebar.style.display = tab === 'meeting' ? '' : 'none';
        mainContent.style.marginLeft = tab === 'meeting' ? '' : '0';
      }
    });
  });

  // Scroll container is .app-layout, not the window
  const scrollContainer = document.querySelector<HTMLElement>('.app-layout')!;

  // Focus tracking: clicking a section card focuses it
  function setFocusedSection(num: number) {
    document.querySelectorAll<HTMLElement>('.section-card').forEach(card => {
      card.classList.toggle('focused', card.id === `sec-${num}`);
    });
    document.querySelectorAll<HTMLAnchorElement>('.sidebar-item').forEach(item => {
      item.classList.toggle('active', item.dataset.nav === String(num));
    });
  }

  // Click anywhere on a section card to focus it
  document.querySelectorAll<HTMLElement>('.section-card[id^="sec-"]').forEach(card => {
    card.addEventListener('click', () => {
      const num = parseInt(card.id.replace('sec-', ''));
      setFocusedSection(num);
    });
  });

  // Sidebar nav clicks: focus + scroll to section
  document.querySelectorAll<HTMLAnchorElement>('.sidebar-item[data-nav]').forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      const n = parseInt(link.dataset.nav!);
      setFocusedSection(n);
      const el = document.getElementById(`sec-${n}`);
      if (el && scrollContainer) {
        const top = el.getBoundingClientRect().top + scrollContainer.scrollTop - scrollContainer.getBoundingClientRect().top;
        scrollContainer.scrollTo({ top, behavior: 'smooth' });
      }
    });
  });

  // Update sidebar on scroll based on which section is most visible
  let _scrollTimer: ReturnType<typeof setTimeout> | null = null;
  scrollContainer.addEventListener('scroll', () => {
    if (_scrollTimer) clearTimeout(_scrollTimer);
    _scrollTimer = setTimeout(() => {
      const cards = document.querySelectorAll<HTMLElement>('.section-card[id^="sec-"]');
      const containerTop = scrollContainer.getBoundingClientRect().top;
      const containerMid = containerTop + scrollContainer.clientHeight / 3;
      let closest: number | null = null;
      let closestDist = Infinity;
      cards.forEach(card => {
        const rect = card.getBoundingClientRect();
        const dist = Math.abs(rect.top - containerMid);
        if (dist < closestDist) {
          closestDist = dist;
          closest = parseInt(card.id.replace('sec-', ''));
        }
      });
      if (closest !== null) setFocusedSection(closest);
    }, 50);
  });

  // Default focus on section 1
  setFocusedSection(1);

  // Section collapse — only title text and chevron trigger it
  document.querySelectorAll<HTMLElement>('[data-section] h2').forEach(h2 => {
    h2.style.cursor = 'pointer';
    h2.addEventListener('click', (e) => {
      e.stopPropagation();
      const n = h2.closest('[data-section]')!.getAttribute('data-section')!;
      document.getElementById(`body-${n}`)?.classList.toggle('collapsed');
      document.getElementById(`chev-${n}`)?.classList.toggle('open');
    });
  });

  // Timer play/pause
  document.querySelectorAll<HTMLButtonElement>('[data-timer]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleTimer(parseInt(btn.dataset.timer!));
    });
  });

  document.querySelectorAll<HTMLButtonElement>('[data-timer-reset]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      resetTimer(parseInt(btn.dataset.timerReset!));
    });
  });

  // Add row buttons
  function addOrToast(container: string, adder: () => void): void {
    const before = document.querySelectorAll(container).length;
    adder();
    if (document.querySelectorAll(container).length === before) showCapToast();
  }
  document.getElementById('btnAddScorecard')?.addEventListener('click', () => addOrToast('#scorecardTable tbody tr', addScorecardRow));
  document.getElementById('btnAddOkrReview')?.addEventListener('click', () => addOrToast('#okrReviewTable tbody tr', addOkrReviewRow));
  document.getElementById('btnAddHeadline')?.addEventListener('click', () => addOrToast('#headlinesTable tbody tr', addHeadlineRow));
  document.getElementById('btnAddTodoReview')?.addEventListener('click', () => addOrToast('#todoReviewTable tbody tr', addTodoReviewRow));
  document.getElementById('btnAddIssue')?.addEventListener('click', () => addOrToast('#issuesListTable tbody tr', addIssueRow));
  document.getElementById('btnAddIDSIssue')?.addEventListener('click', () => addOrToast('#idsIssuesContainer .ids-issue', addIDSIssue));
  document.getElementById('btnAddNewTodo')?.addEventListener('click', () => addOrToast('#newTodoTable tbody tr', addNewTodoRow));
  document.getElementById('btnAddCascading')?.addEventListener('click', () => addOrToast('#cascadingTable tbody tr', addCascadingRow));
  document.getElementById('btnAddRating')?.addEventListener('click', () => addOrToast('#ratingTable tbody tr', addRatingRow));
  document.getElementById('btnAddScorecardFull')?.addEventListener('click', () => addOrToast('#scorecardFullTable tbody tr', addScorecardFullRow));
  document.getElementById('btnAddOkrFull')?.addEventListener('click', () => addOrToast('#okrFullTable tbody tr', addOkrFullRow));
}

// ── Standalone one-time meeting ──

function initStandaloneMeeting(): void {
  const app = document.getElementById('app')!;
  resetIdsIssueCount();
  app.innerHTML = buildAppHTML('', true);

  // Show top bar actions immediately (back/export buttons)
  const actions = document.getElementById('topBarActions')!;
  actions.style.opacity = '1';
  actions.style.pointerEvents = '';

  // Auto-fill date
  const now = new Date();
  (document.getElementById('metaDate') as HTMLInputElement).value = now.toISOString().split('T')[0];

  // Duration helper
  function updateDuration(): void {
    const controlDiv = document.querySelector('.meeting-control');
    if (!controlDiv) return;
    const startVal = (document.getElementById('metaStart') as HTMLInputElement)?.value;
    const endVal = (document.getElementById('metaEnd') as HTMLInputElement)?.value;
    if (!startVal || !endVal) {
      controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">&nbsp;</span>`;
      return;
    }
    const sp = startVal.split(':').map(Number);
    const ep = endVal.split(':').map(Number);
    if (sp.length < 2 || ep.length < 2 || sp.some(isNaN) || ep.some(isNaN)) {
      controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">&nbsp;</span>`;
      return;
    }
    const totalMins = (ep[0] * 60 + ep[1]) - (sp[0] * 60 + sp[1]);
    if (totalMins <= 0) {
      controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">&nbsp;</span>`;
      return;
    }
    const h = Math.floor(totalMins / 60);
    const m = totalMins % 60;
    const elapsed = h > 0 ? `${h}:${String(m).padStart(2, '0')}:00` : `${m}:00`;
    controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">Duration: ${elapsed}</span>`;
  }

  // Meeting start/stop
  const meetingTab = document.getElementById('tab-meeting')!;
  const sidebar = document.getElementById('sidebar')!;
  meetingTab.classList.add('blurred');
  sidebar.classList.add('blurred');

  let meetingInterval: ReturnType<typeof setInterval> | null = null;
  let meetingSeconds = 0;

  // Push hash so browser back returns to folder picker
  history.pushState(null, '', '#/onetime');
  function goBackToFolderPicker() {
    window.removeEventListener('hashchange', onStandaloneHash);
    if (meetingInterval) clearInterval(meetingInterval);
    cleanupTimers();
    fs.hasStoredFolder().then(stored => showFolderPicker(stored === 'prompt'));
  }
  const onStandaloneHash = () => {
    if (location.hash !== '#/onetime') goBackToFolderPicker();
  };
  window.addEventListener('hashchange', onStandaloneHash);

  function formatElapsed(secs: number): string {
    const h = Math.floor(secs / 3600);
    const m = Math.floor((secs % 3600) / 60);
    const s = secs % 60;
    if (h > 0) return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
    return `${m}:${String(s).padStart(2, '0')}`;
  }

  function startMeeting() {
    meetingTab.classList.remove('blurred');
    sidebar.classList.remove('blurred');
    const actions = document.getElementById('topBarActions')!;
    actions.style.opacity = '1';
    actions.style.pointerEvents = '';
    const startNow = new Date();
    (document.getElementById('metaStart') as HTMLInputElement).value = startNow.toTimeString().slice(0, 5);

    const controlDiv = document.querySelector('.meeting-control')!;
    meetingSeconds = 0;
    controlDiv.innerHTML = `
      <button class="meeting-stop-btn" id="btnMeetingStop">
        <span class="stop-icon"></span>
        <span class="meeting-timer-display" id="meetingElapsed">0:00</span>
      </button>`;

    meetingInterval = setInterval(() => {
      meetingSeconds++;
      const display = document.getElementById('meetingElapsed');
      if (display) display.textContent = formatElapsed(meetingSeconds);
    }, 1000);

    document.getElementById('btnMeetingStop')!.addEventListener('click', stopMeeting);
  }

  function stopMeeting() {
    if (meetingInterval) clearInterval(meetingInterval);
    const endNow = new Date();
    (document.getElementById('metaEnd') as HTMLInputElement).value = endNow.toTimeString().slice(0, 5);
    updateDuration();
  }

  document.getElementById('btnMeetingStart')!.addEventListener('click', startMeeting);

  // Init timers
  initTimers();

  // Populate default rows
  DEFAULT_MEASURABLES.forEach(m => addScorecardRow(m));
  for (let i = 0; i < DEFAULT_ROWS.okr; i++) addOkrReviewRow();
  for (let i = 0; i < DEFAULT_ROWS.headlines; i++) addHeadlineRow();
  for (let i = 0; i < DEFAULT_ROWS.todoReview; i++) addTodoReviewRow();
  for (let i = 0; i < DEFAULT_ROWS.issues; i++) addIssueRow();
  for (let i = 0; i < DEFAULT_ROWS.idsIssues; i++) addIDSIssue();
  for (let i = 0; i < DEFAULT_ROWS.newTodos; i++) addNewTodoRow();
  for (let i = 0; i < DEFAULT_ROWS.cascading; i++) addCascadingRow();
  for (let i = 0; i < DEFAULT_ROWS.rating; i++) addRatingRow();

  DEFAULT_MEASURABLES.forEach(m => addScorecardFullRow(m));
  for (let i = 1; i <= DEFAULT_ROWS.okr; i++) addOkrFullRow('', i);
  buildKeyResultBlocks();

  // Recalc duration on time changes
  document.getElementById('metaStart')?.addEventListener('change', updateDuration);
  document.getElementById('metaEnd')?.addEventListener('change', updateDuration);

  // Export Excel button
  document.getElementById('btnExportExcel')?.addEventListener('click', async () => {
    const { gatherMeetingData } = await import('./storage');
    const data = gatherMeetingData();
    const buffer = await fs.exportMeetingToBuffer(data);
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const date = (document.getElementById('metaDate') as HTMLInputElement)?.value || new Date().toISOString().split('T')[0];
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `L10_Meeting_${date}.xlsx`;
    a.click();
    URL.revokeObjectURL(a.href);
  });

  // Back button → return to folder picker
  document.getElementById('btnStandaloneBack')?.addEventListener('click', () => {
    history.back();
  });

  // Tab switching
  document.querySelectorAll<HTMLButtonElement>('.top-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab!;
      document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.top-tab').forEach(b => b.classList.remove('active'));
      document.getElementById(`tab-${tab}`)?.classList.add('active');
      btn.classList.add('active');
      const mainContent = document.querySelector<HTMLElement>('.main-content');
      if (mainContent) {
        sidebar.style.display = tab === 'meeting' ? '' : 'none';
        mainContent.style.marginLeft = tab === 'meeting' ? '' : '0';
      }
    });
  });

  // Scroll container
  const scrollContainer = document.querySelector<HTMLElement>('.app-layout')!;

  function setFocusedSection(num: number) {
    document.querySelectorAll<HTMLElement>('.section-card').forEach(card => {
      card.classList.toggle('focused', card.id === `sec-${num}`);
    });
    document.querySelectorAll<HTMLAnchorElement>('.sidebar-item').forEach(item => {
      item.classList.toggle('active', item.dataset.nav === String(num));
    });
  }

  document.querySelectorAll<HTMLElement>('.section-card[id^="sec-"]').forEach(card => {
    card.addEventListener('click', () => {
      const num = parseInt(card.id.replace('sec-', ''));
      setFocusedSection(num);
    });
  });

  document.querySelectorAll<HTMLAnchorElement>('.sidebar-item[data-nav]').forEach(link => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      const n = parseInt(link.dataset.nav!);
      setFocusedSection(n);
      const el = document.getElementById(`sec-${n}`);
      if (el && scrollContainer) {
        const top = el.getBoundingClientRect().top + scrollContainer.scrollTop - scrollContainer.getBoundingClientRect().top;
        scrollContainer.scrollTo({ top, behavior: 'smooth' });
      }
    });
  });

  let _scrollTimer: ReturnType<typeof setTimeout> | null = null;
  scrollContainer.addEventListener('scroll', () => {
    if (_scrollTimer) clearTimeout(_scrollTimer);
    _scrollTimer = setTimeout(() => {
      const cards = document.querySelectorAll<HTMLElement>('.section-card[id^="sec-"]');
      const containerTop = scrollContainer.getBoundingClientRect().top;
      const containerMid = containerTop + scrollContainer.clientHeight / 3;
      let closest: number | null = null;
      let closestDist = Infinity;
      cards.forEach(card => {
        const rect = card.getBoundingClientRect();
        const dist = Math.abs(rect.top - containerMid);
        if (dist < closestDist) { closestDist = dist; closest = parseInt(card.id.replace('sec-', '')); }
      });
      if (closest !== null) setFocusedSection(closest);
    }, 50);
  });

  setFocusedSection(1);

  // Section collapse
  document.querySelectorAll<HTMLElement>('[data-section] h2').forEach(h2 => {
    h2.style.cursor = 'pointer';
    h2.addEventListener('click', (e) => {
      e.stopPropagation();
      const n = h2.closest('[data-section]')!.getAttribute('data-section')!;
      document.getElementById(`body-${n}`)?.classList.toggle('collapsed');
      document.getElementById(`chev-${n}`)?.classList.toggle('open');
    });
  });

  // Timer play/pause
  document.querySelectorAll<HTMLButtonElement>('[data-timer]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleTimer(parseInt(btn.dataset.timer!));
    });
  });
  document.querySelectorAll<HTMLButtonElement>('[data-timer-reset]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      resetTimer(parseInt(btn.dataset.timerReset!));
    });
  });

  // Add row buttons
  function addOrToast(container: string, adder: () => void): void {
    const before = document.querySelectorAll(container).length;
    adder();
    if (document.querySelectorAll(container).length === before) showCapToast();
  }
  document.getElementById('btnAddScorecard')?.addEventListener('click', () => addOrToast('#scorecardTable tbody tr', addScorecardRow));
  document.getElementById('btnAddOkrReview')?.addEventListener('click', () => addOrToast('#okrReviewTable tbody tr', addOkrReviewRow));
  document.getElementById('btnAddHeadline')?.addEventListener('click', () => addOrToast('#headlinesTable tbody tr', addHeadlineRow));
  document.getElementById('btnAddTodoReview')?.addEventListener('click', () => addOrToast('#todoReviewTable tbody tr', addTodoReviewRow));
  document.getElementById('btnAddIssue')?.addEventListener('click', () => addOrToast('#issuesListTable tbody tr', addIssueRow));
  document.getElementById('btnAddIDSIssue')?.addEventListener('click', () => addOrToast('#idsIssuesContainer .ids-issue', addIDSIssue));
  document.getElementById('btnAddNewTodo')?.addEventListener('click', () => addOrToast('#newTodoTable tbody tr', addNewTodoRow));
  document.getElementById('btnAddCascading')?.addEventListener('click', () => addOrToast('#cascadingTable tbody tr', addCascadingRow));
  document.getElementById('btnAddRating')?.addEventListener('click', () => addOrToast('#ratingTable tbody tr', addRatingRow));
  document.getElementById('btnAddScorecardFull')?.addEventListener('click', () => addOrToast('#scorecardFullTable tbody tr', addScorecardFullRow));
  document.getElementById('btnAddOkrFull')?.addEventListener('click', () => addOrToast('#okrFullTable tbody tr', addOkrFullRow));

  // Person pickers
  initPersonPickers();
}

// ── Folder picker landing page ──

function showFolderPicker(hasStored: boolean): void {
  const app = document.getElementById('app')!;
  app.innerHTML = `
    <div class="fp-bg">
      <div class="fp-glow fp-glow-1"></div>
      <div class="fp-glow fp-glow-2"></div>
    </div>
    <div class="folder-picker">
      <div class="fp-card">
        <h1 class="fp-title">L10 Meeting Manager</h1>
        <p class="fp-desc">Select a new or existing folder to store your meeting data. This choice is remembered.</p>

        <div class="fp-divider"></div>

        <p class="fp-step-label">${hasStored ? 'Reconnect to your data' : 'Get started'}</p>

        <button class="fp-btn fp-btn-dark" id="btnOneTime">One-Time Meeting</button>

        ${hasStored ? `
          <button class="fp-btn fp-btn-primary" id="btnRestoreFolder">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
            Reconnect folder
          </button>
          <button class="fp-btn fp-btn-ghost" id="btnPickFolder">Choose a different folder</button>
          <button class="fp-btn fp-btn-ghost fp-btn-danger" id="btnForgetFolder">Forget saved folder</button>
        ` : `
          <button class="fp-btn fp-btn-primary" id="btnPickFolder">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
            Choose a folder
          </button>
        `}

        <div class="fp-tip">
          <svg class="fp-tip-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18h6M10 22h4M12 2a7 7 0 0 0-4 12.7V17h8v-2.3A7 7 0 0 0 12 2z"/></svg>
          Tip: Use a folder synced with OneDrive, Google Drive, or Dropbox for backup and multi-device access.
        </div>
      </div>
      <div class="fp-tip fp-tip-green">
        <svg class="fp-tip-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
        No data is collected. All your data stays locally on your machine.
      </div>
    </div>
    <div class="toast"></div>
  `;

  if (hasStored) {
    document.getElementById('btnRestoreFolder')?.addEventListener('click', async () => {
      const ok = await fs.restoreFolder();
      if (ok) startApp();
      else alert('Permission denied. Please try again or choose a new folder.');
    });
    document.getElementById('btnForgetFolder')?.addEventListener('click', async () => {
      await fs.forgetFolder();
      showFolderPicker(false);
    });
  }

  document.getElementById('btnPickFolder')?.addEventListener('click', async () => {
    if (!('showDirectoryPicker' in window)) {
      const t = document.querySelector<HTMLElement>('.toast');
      if (t) {
        t.textContent = 'Folder access is not supported on this device';
        t.classList.add('show', 'toast-error');
        setTimeout(() => t.classList.remove('show', 'toast-error'), 3000);
      }
      return;
    }
    const ok = await fs.pickFolder();
    if (ok) startApp();
  });

  document.getElementById('btnOneTime')?.addEventListener('click', () => {
    initStandaloneMeeting();
  });
}

async function startApp(): Promise<void> {
  await initLogo();
  initPersonPickers();
  window.addEventListener('hashchange', route);
  route();
}

// ── Boot ──
(async () => {
  const stored = await fs.hasStoredFolder();
  if (stored === 'granted') {
    await fs.restoreFolder();
    await startApp();
  } else {
    showFolderPicker(stored === 'prompt');
  }
})();
