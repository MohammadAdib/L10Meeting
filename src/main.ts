import './style.css';
import { buildAppHTML } from './html';
import { initTimers, toggleTimer, resetTimer } from './timer';
import { onStatusChange } from './utils';
import { resetAll, isServerAvailable, saveMeeting, listMeetings, loadMeeting } from './storage';
import { exportExcel } from './export';
import { DEFAULT_MEASURABLES } from './types';
import {
  addScorecardRow, addOkrReviewRow, addHeadlineRow, addTodoReviewRow,
  addIssueRow, addIDSIssue, addIDSTodoRow, addNewTodoRow, addCascadingRow,
  addRatingRow, setRating, updateTodoCompletion, updateAvgRating,
  addScorecardFullRow, addOkrFullRow, addKeyResultRow, buildKeyResultBlocks,
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

// ── Render ──
document.querySelector<HTMLDivElement>('#app')!.innerHTML = buildAppHTML();

// ── Auto-fill date ──
const now = new Date();
(document.getElementById('metaDate') as HTMLInputElement).value = now.toISOString().split('T')[0];

// ── Meeting start/stop ──
const meetingTab = document.getElementById('tab-meeting')!;
const sidebar = document.getElementById('sidebar')!;
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
  meetingTab.classList.remove('blurred');
  sidebar.classList.remove('blurred');
  const actions = document.getElementById('topBarActions')!;
  actions.style.opacity = '1';
  actions.style.pointerEvents = '';
  // Fill start time
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
  // Fill end time
  const endNow = new Date();
  (document.getElementById('metaEnd') as HTMLInputElement).value = endNow.toTimeString().slice(0, 5);

  const controlDiv = document.querySelector('.meeting-control')!;
  controlDiv.innerHTML = `<span style="color:var(--text-muted);font-size:13px;font-weight:600;">Meeting ended — ${formatElapsed(meetingSeconds)}</span>`;
}

document.getElementById('btnMeetingStart')!.addEventListener('click', startMeeting);

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
for (let i = 0; i < 5; i++) addRatingRow();
DEFAULT_MEASURABLES.concat(['', '', '']).forEach(m => addScorecardFullRow(m));
for (let i = 1; i <= 7; i++) addOkrFullRow('', i);
buildKeyResultBlocks();

// ── Event Delegation ──

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
    if (el) {
      const top = el.getBoundingClientRect().top + scrollContainer.scrollTop - scrollContainer.getBoundingClientRect().top;
      scrollContainer.scrollTo({ top, behavior: 'smooth' });
    }
  });
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
document.getElementById('btnSave')?.addEventListener('click', saveMeeting);
document.getElementById('btnReset')?.addEventListener('click', resetAll);
document.getElementById('btnExport')?.addEventListener('click', exportExcel);

// Load dropdown
const loadMenuBtn = document.getElementById('btnLoadMenu');
const loadDropdown = document.getElementById('loadDropdown');
loadMenuBtn?.addEventListener('click', async () => {
  const meetings = await listMeetings();
  if (meetings.length === 0) {
    loadDropdown!.innerHTML = '<div class="load-item empty">No saved meetings</div>';
  } else {
    loadDropdown!.innerHTML = meetings.map(m =>
      `<div class="load-item" data-file="${m.filename}">${m.filename.replace('.json', '').replace(/_/g, ' ')}</div>`
    ).join('');
    loadDropdown!.querySelectorAll<HTMLElement>('.load-item[data-file]').forEach(item => {
      item.addEventListener('click', () => {
        loadMeeting(item.dataset.file!);
        loadDropdown!.classList.remove('open');
      });
    });
  }
  loadDropdown!.classList.toggle('open');
});
document.addEventListener('click', (e) => {
  if (!loadMenuBtn?.contains(e.target as Node) && !loadDropdown?.contains(e.target as Node)) {
    loadDropdown?.classList.remove('open');
  }
});

// Detect server and show save/load buttons
isServerAvailable().then(available => {
  if (available) {
    document.querySelectorAll<HTMLElement>('.server-only').forEach(el => el.style.display = '');
  }
});

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
