import './style.css';
import { buildAppHTML } from './html';
import { initTimers, toggleTimer, resetTimer } from './timer';
import { onStatusChange } from './utils';
import { resetAll } from './storage';
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
const appLayout = document.querySelector<HTMLElement>('.app-layout')!;
appLayout.classList.add('blurred');
document.body.classList.add('no-scroll');

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
  appLayout.classList.remove('blurred');
  document.body.classList.remove('no-scroll');
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
    const sidebar = document.getElementById('sidebar');
    const mainContent = document.querySelector<HTMLElement>('.main-content');
    if (sidebar && mainContent) {
      sidebar.style.display = tab === 'meeting' ? '' : 'none';
      mainContent.style.marginLeft = tab === 'meeting' ? '' : '0';
    }
  });
});

// Sidebar nav clicks
document.querySelectorAll<HTMLAnchorElement>('.sidebar-item[data-nav]').forEach(link => {
  link.addEventListener('click', (e) => {
    e.preventDefault();
    const n = link.dataset.nav!;
    const el = document.getElementById(`sec-${n}`);
    if (el) {
      const top = el.getBoundingClientRect().top + window.scrollY - 60;
      window.scrollTo({ top, behavior: 'smooth' });
    }
  });
});

// Scroll-spy: highlight current section in sidebar
function updateScrollSpy() {
  const sections = [1, 2, 3, 4, 5, 6, 7];
  let current = 1;
  for (const n of sections) {
    const el = document.getElementById(`sec-${n}`);
    if (el) {
      const rect = el.getBoundingClientRect();
      if (rect.top <= 120) current = n;
    }
  }
  document.querySelectorAll<HTMLAnchorElement>('.sidebar-item').forEach(item => {
    item.classList.toggle('active', item.dataset.nav === String(current));
  });
}
window.addEventListener('scroll', updateScrollSpy, { passive: true });
updateScrollSpy();

// Section collapse
document.querySelectorAll<HTMLDivElement>('[data-section]').forEach(header => {
  header.addEventListener('click', () => {
    const n = header.dataset.section!;
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
document.getElementById('btnExport')?.addEventListener('click', exportExcel);

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
