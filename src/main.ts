import './style.css';
import { buildAppHTML } from './html';
import { initTimers, toggleTimer, resetTimer, cleanupTimers } from './timer';
import { onStatusChange } from './utils';
import { resetAll, loadMeetingData, setupAutoSave, markMeetingStarted, isMeetingActive, cleanupAutoSave, openInExcel } from './storage';
import { DEFAULT_MEASURABLES } from './types';
import { renderAdminPortal, renderDepartmentView } from './admin';
import {
  addScorecardRow, addOkrReviewRow, addHeadlineRow, addTodoReviewRow,
  addIssueRow, addIDSIssue, addIDSTodoRow, addNewTodoRow, addCascadingRow,
  addRatingRow, setRating, updateTodoCompletion, updateAvgRating,
  addScorecardFullRow, addOkrFullRow, addKeyResultRow, buildKeyResultBlocks,
  resetIdsIssueCount,
} from './tables';

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
window.__addIDSTodoRow = addIDSTodoRow;
window.__addKeyResultRow = addKeyResultRow;

// ── Router ──
let _previousHash = '';

async function route() {
  const hash = location.hash || '#/';
  const leavingMeeting = _previousHash.includes('/meeting/') && !hash.includes('/meeting/');

  // Confirm before leaving an active meeting
  if (leavingMeeting && isMeetingActive()) {
    if (!confirm('You have an active meeting. Are you sure you want to leave?')) {
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
    controlDiv.innerHTML = '';
    document.querySelectorAll<HTMLElement>('.section-timer').forEach(el => el.style.display = 'none');
    document.getElementById('btnReset')?.remove();
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
      const endNow = new Date();
      (document.getElementById('metaEnd') as HTMLInputElement).value = endNow.toTimeString().slice(0, 5);

      const controlDiv = document.querySelector('.meeting-control')!;
      controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">Meeting ended — ${formatElapsed(meetingSeconds)}</span>`;
    }

    document.getElementById('btnMeetingStart')!.addEventListener('click', startMeeting);
  }

  // ── Init timers ──
  initTimers();

  // ── Populate default rows ──
  DEFAULT_MEASURABLES.forEach(m => addScorecardRow(m));
  for (let i = 0; i < 6; i++) addOkrReviewRow();
  for (let i = 0; i < 5; i++) addHeadlineRow();
  for (let i = 0; i < 7; i++) addTodoReviewRow();
  for (let i = 0; i < 5; i++) addIssueRow();
  for (let i = 0; i < 3; i++) addIDSIssue();
  for (let i = 0; i < 7; i++) addNewTodoRow();
  for (let i = 0; i < 3; i++) addCascadingRow();

  // ── Pre-fill rating table with department people ──
  let people: string[] = [];
  try {
    const res = await fetch(`/api/departments/${encodeURIComponent(deptName)}/people`);
    if (res.ok) people = await res.json();
  } catch { /* empty */ }

  if (people.length > 0) {
    people.forEach(() => addRatingRow());
    // Fill in names
    const nameInputs = document.querySelectorAll<HTMLInputElement>('#ratingTable tbody tr input[placeholder="Name"]');
    people.forEach((name, i) => {
      if (i < nameInputs.length) nameInputs[i].value = name;
    });
  } else {
    for (let i = 0; i < 5; i++) addRatingRow();
  }

  DEFAULT_MEASURABLES.concat(['', '', '']).forEach(m => addScorecardFullRow(m));
  for (let i = 1; i <= 7; i++) addOkrFullRow('', i);
  buildKeyResultBlocks();

  // ── If loading existing meeting, populate data ──
  if (meetingId !== 'new') {
    try {
      const res = await fetch(`/api/departments/${encodeURIComponent(deptName)}/meetings/${meetingId}`);
      if (res.ok) {
        const data = await res.json();
        loadMeetingData(data);
      }
    } catch { /* new meeting */ }
  }

  // ── Set up auto-save ──
  // For new meetings, defer creating on server until first save
  setupAutoSave(deptName, meetingId === 'new' ? '' : meetingId, meetingId === 'new');

  // ── Event Delegation ──

  // Logo click → back to department view
  document.querySelector('.top-bar-logo')?.addEventListener('click', () => {
    location.hash = `#/dept/${encodeURIComponent(deptName)}`;
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

  // Timer reset
  document.querySelectorAll<HTMLButtonElement>('[data-timer-reset]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      resetTimer(parseInt(btn.dataset.timerReset!));
    });
  });

  // Top bar buttons
  document.getElementById('btnReset')?.addEventListener('click', resetAll);
  document.getElementById('btnOpenExcel')?.addEventListener('click', openInExcel);

  // Add row buttons
  document.getElementById('btnAddScorecard')?.addEventListener('click', () => addScorecardRow());
  document.getElementById('btnAddOkrReview')?.addEventListener('click', () => addOkrReviewRow());
  document.getElementById('btnAddHeadline')?.addEventListener('click', () => addHeadlineRow());
  document.getElementById('btnAddTodoReview')?.addEventListener('click', () => addTodoReviewRow());
  document.getElementById('btnAddIssue')?.addEventListener('click', () => addIssueRow());
  document.getElementById('btnAddIDSIssue')?.addEventListener('click', () => addIDSIssue());
  document.getElementById('btnAddNewTodo')?.addEventListener('click', () => addNewTodoRow());
  document.getElementById('btnAddCascading')?.addEventListener('click', () => addCascadingRow());
  document.getElementById('btnAddRating')?.addEventListener('click', () => addRatingRow());
  document.getElementById('btnAddScorecardFull')?.addEventListener('click', () => addScorecardFullRow());
  document.getElementById('btnAddOkrFull')?.addEventListener('click', () => addOkrFullRow());
}

// ── Start the router ──
window.addEventListener('hashchange', route);
route();
