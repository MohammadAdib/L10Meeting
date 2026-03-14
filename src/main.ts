import './style.css';
import { buildAppHTML } from './html';
import { initTimers, toggleTimer, resetTimer } from './timer';
import { onStatusChange } from './utils';
import { saveDraft, loadDraft, resetAll } from './storage';
import { exportPDF } from './pdf';
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

// ── Auto-fill date/time ──
const now = new Date();
(document.getElementById('metaDate') as HTMLInputElement).value = now.toISOString().split('T')[0];
(document.getElementById('metaStart') as HTMLInputElement).value = now.toTimeString().slice(0, 5);
const end = new Date(now.getTime() + 90 * 60000);
(document.getElementById('metaEnd') as HTMLInputElement).value = end.toTimeString().slice(0, 5);

// ── Init timers ──
initTimers();

// ── Populate default rows ──
DEFAULT_MEASURABLES.forEach(m => addScorecardRow(m));
for (let i = 1; i <= 6; i++) addOkrReviewRow(`OKR ${i}`);
for (let i = 0; i < 5; i++) addHeadlineRow();
for (let i = 0; i < 7; i++) addTodoReviewRow();
for (let i = 0; i < 5; i++) addIssueRow();
for (let i = 0; i < 3; i++) addIDSIssue();
for (let i = 0; i < 7; i++) addNewTodoRow();
for (let i = 0; i < 3; i++) addCascadingRow();
for (let i = 0; i < 5; i++) addRatingRow();
DEFAULT_MEASURABLES.concat(['', '', '']).forEach(m => addScorecardFullRow(m));
for (let i = 1; i <= 7; i++) addOkrFullRow(i <= 6 ? `OKR ${i}` : '', i);
buildKeyResultBlocks();

// ── Event Delegation ──

// Tab switching
document.querySelectorAll<HTMLButtonElement>('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const tab = btn.dataset.tab!;
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById(`tab-${tab}`)?.classList.add('active');
    btn.classList.add('active');
  });
});

// Section nav
document.querySelectorAll<HTMLButtonElement>('[data-nav]').forEach(btn => {
  btn.addEventListener('click', () => {
    const n = btn.dataset.nav!;
    document.getElementById(`sec-${n}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    document.querySelectorAll('.section-nav button').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  });
});

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
document.getElementById('btnSave')?.addEventListener('click', saveDraft);
document.getElementById('btnLoad')?.addEventListener('click', loadDraft);
document.getElementById('btnReset')?.addEventListener('click', resetAll);
document.getElementById('btnExport')?.addEventListener('click', exportPDF);

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
