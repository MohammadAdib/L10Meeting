import { showToast, getTableRows, val, populateTableRows } from './utils';
import {
  updateTodoCompletion, updateAvgRating,
  addScorecardRow, addOkrReviewRow, addHeadlineRow, addTodoReviewRow,
  addIssueRow, addIDSIssue, addIDSTodoRow, addNewTodoRow, addCascadingRow,
  addRatingRow, addScorecardFullRow, addOkrFullRow,
} from './tables';
import { createMeeting, saveMeeting } from './fs-service';

/** Gather all meeting data from the DOM into a JSON-serializable object */
export function gatherMeetingData(): Record<string, unknown> {
  const data: Record<string, unknown> = {
    meta: {
      team: val('metaTeam'),
      date: val('metaDate'),
      facilitator: val('metaFacilitator'),
      scribe: val('metaScribe'),
      start: val('metaStart'),
      end: val('metaEnd'),
    },
    segue: {
      personal: (document.getElementById('seguePersonal') as HTMLTextAreaElement)?.value ?? '',
      professional: (document.getElementById('segueProfessional') as HTMLTextAreaElement)?.value ?? '',
    },
    okrMeta: {
      quarter: (document.getElementById('okrQuarter') as HTMLSelectElement)?.value ?? '',
      year: (document.getElementById('okrYear') as HTMLSelectElement)?.value ?? '',
      startDate: (document.getElementById('okrStartDate') as HTMLInputElement)?.value ?? '',
      targetDate: (document.getElementById('okrTargetDate') as HTMLInputElement)?.value ?? '',
    },
  };

  // Gather all table data (including scorecard/OKR full tabs)
  const tableIds = [
    'scorecardTable', 'okrReviewTable', 'headlinesTable', 'todoReviewTable',
    'issuesListTable', 'newTodoTable', 'cascadingTable', 'ratingTable',
    'scorecardFullTable', 'okrFullTable',
  ];
  for (const id of tableIds) {
    data[id] = getTableRows(id);
  }

  // IDS issue detail blocks
  const idsBlocks: { fields: string[]; todos: string[][] }[] = [];
  document.querySelectorAll('#idsIssuesContainer .ids-issue').forEach((block, i) => {
    const fields: string[] = [];
    block.querySelectorAll<HTMLTextAreaElement>('.ids-field textarea').forEach(ta => fields.push(ta.value));
    idsBlocks.push({ fields, todos: getTableRows(`idsTodo-${i + 1}`) });
  });
  data.idsBlocks = idsBlocks;

  // Key result blocks
  const keyResults: string[][][] = [];
  document.querySelectorAll<HTMLTableElement>('[id^="keyResults-"]').forEach(table => {
    keyResults.push(getTableRows(table.id));
  });
  data.keyResults = keyResults;

  // Rating values (hidden inputs)
  const ratingValues: string[] = [];
  document.querySelectorAll<HTMLInputElement>('#ratingTable .rating-value').forEach(el => {
    ratingValues.push(el.value);
  });
  data.ratingValues = ratingValues;

  return data;
}

/** Populate DOM from saved meeting data */
export function loadMeetingData(data: Record<string, unknown>): void {
  const meta = data.meta as Record<string, string> | undefined;
  if (meta) {
    const map: Record<string, string> = {
      team: 'metaTeam', date: 'metaDate', facilitator: 'metaFacilitator',
      scribe: 'metaScribe', start: 'metaStart', end: 'metaEnd',
    };
    for (const [key, id] of Object.entries(map)) {
      const el = document.getElementById(id) as HTMLInputElement | null;
      if (el) el.value = meta[key] ?? '';
    }
  }

  const segue = data.segue as Record<string, string> | undefined;
  if (segue) {
    const p = document.getElementById('seguePersonal') as HTMLTextAreaElement | null;
    const pr = document.getElementById('segueProfessional') as HTMLTextAreaElement | null;
    if (p) p.value = segue.personal ?? '';
    if (pr) pr.value = segue.professional ?? '';
  }

  const okrMeta = data.okrMeta as Record<string, string> | undefined;
  if (okrMeta) {
    const map: Record<string, string> = {
      quarter: 'okrQuarter', year: 'okrYear',
      startDate: 'okrStartDate', targetDate: 'okrTargetDate',
    };
    for (const [key, id] of Object.entries(map)) {
      const el = document.getElementById(id) as HTMLInputElement | HTMLSelectElement | null;
      if (el) el.value = okrMeta[key] ?? '';
    }
  }

  // Restore tables — add extra rows if data has more than pre-created
  const tableAdders: Record<string, () => void> = {
    scorecardTable: addScorecardRow,
    okrReviewTable: addOkrReviewRow,
    headlinesTable: addHeadlineRow,
    todoReviewTable: addTodoReviewRow,
    issuesListTable: addIssueRow,
    newTodoTable: addNewTodoRow,
    cascadingTable: addCascadingRow,
    ratingTable: addRatingRow,
    scorecardFullTable: addScorecardFullRow,
    okrFullTable: addOkrFullRow,
  };
  const tableIds = [
    'scorecardTable', 'okrReviewTable', 'headlinesTable', 'todoReviewTable',
    'issuesListTable', 'newTodoTable', 'cascadingTable', 'ratingTable',
    'scorecardFullTable', 'okrFullTable',
  ];
  for (const tableId of tableIds) {
    const rows = data[tableId] as string[][] | undefined;
    if (!rows) continue;
    const adder = tableAdders[tableId];
    if (adder) {
      const existing = document.querySelectorAll(`#${tableId} tbody tr`).length;
      for (let i = existing; i < rows.length; i++) adder();
    }
    populateTableRows(`#${tableId}`, rows);
  }

  // IDS blocks — add extra blocks if needed, then populate fields and todos
  const idsBlocks = data.idsBlocks as { fields: string[]; todos: string[][] }[] | undefined;
  if (idsBlocks) {
    const existingBlocks = document.querySelectorAll('#idsIssuesContainer .ids-issue').length;
    for (let i = existingBlocks; i < idsBlocks.length; i++) addIDSIssue();

    const blocks = document.querySelectorAll('#idsIssuesContainer .ids-issue');
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

  // Key results
  const keyResults = data.keyResults as string[][][] | undefined;
  if (keyResults) {
    keyResults.forEach((rows, ki) => populateTableRows(`#keyResults-${ki + 1}`, rows));
  }

  // Rating values + star visuals
  const ratingValues = data.ratingValues as string[] | undefined;
  if (ratingValues) {
    const inputs = document.querySelectorAll<HTMLInputElement>('#ratingTable .rating-value');
    ratingValues.forEach((v, i) => {
      if (i >= inputs.length) return;
      inputs[i].value = v;
      const val = parseInt(v) || 0;
      const stars = inputs[i].parentElement?.querySelectorAll('.rating-stars button');
      if (stars) stars.forEach((s, si) => s.classList.toggle('active', si < val));
    });
  }

  updateTodoCompletion();
  updateAvgRating();
}

/** Populate scorecard/OKR full tabs and key results from meeting data */
export function loadScorecardOkrData(data: Record<string, unknown>): void {
  // Scorecard full table
  const scRows = data.scorecardFullTable as string[][] | undefined;
  if (scRows) {
    const existing = document.querySelectorAll('#scorecardFullTable tbody tr').length;
    for (let i = existing; i < scRows.length; i++) addScorecardFullRow();
    populateTableRows('#scorecardFullTable', scRows);
  }

  // OKR full table
  const okrRows = data.okrFullTable as string[][] | undefined;
  if (okrRows) {
    const existing = document.querySelectorAll('#okrFullTable tbody tr').length;
    for (let i = existing; i < okrRows.length; i++) addOkrFullRow();
    populateTableRows('#okrFullTable', okrRows);
  }

  // OKR metadata
  const okrMeta = data.okrMeta as Record<string, string> | undefined;
  if (okrMeta) {
    const map: Record<string, string> = {
      quarter: 'okrQuarter', year: 'okrYear',
      startDate: 'okrStartDate', targetDate: 'okrTargetDate',
    };
    for (const [key, id] of Object.entries(map)) {
      const el = document.getElementById(id) as HTMLInputElement | HTMLSelectElement | null;
      if (el && okrMeta[key]) el.value = okrMeta[key];
    }
  }

  // Key results
  const keyResults = data.keyResults as string[][][] | undefined;
  if (keyResults) {
    keyResults.forEach((rows, ki) => populateTableRows(`#keyResults-${ki + 1}`, rows));
  }
}

/** Auto-save: debounced PUT to server */
let _autoSaveTimer: ReturnType<typeof setTimeout> | null = null;
let _autoSaveDept: string = '';
let _autoSaveMeetingId: string = '';
let _meetingDirty = false;
let _meetingStarted = false;
let _isNewMeeting = false;
let _manualSaveMode = false;
let _savedSnapshot: string = '';
let _listenerAbort: AbortController | null = null;

/** Take a JSON snapshot of current meeting state */
function takeSnapshot(): string {
  return JSON.stringify(gatherMeetingData());
}

/** Snapshot the current state as the "clean" baseline */
export function snapshotCleanState(): void {
  _savedSnapshot = takeSnapshot();
  _meetingDirty = false;
}

function checkDirty(): boolean {
  if (!_savedSnapshot) return true; // new meeting with no snapshot = dirty once started
  return takeSnapshot() !== _savedSnapshot;
}

export function markMeetingStarted(): void {
  _meetingStarted = true;
  _meetingDirty = true;
}

export function markMeetingStopped(): void {
  _meetingStarted = false;
}

export function isMeetingActive(): boolean {
  return _meetingStarted;
}

export function isMeetingDirty(): boolean {
  return _meetingDirty;
}

export function setupAutoSave(dept: string, meetingId: string, isNew: boolean = false, manualSave: boolean = false): void {
  // Abort any listeners from a previous setupAutoSave call
  if (_listenerAbort) _listenerAbort.abort();
  _listenerAbort = new AbortController();
  const { signal } = _listenerAbort;

  _autoSaveDept = dept;
  _autoSaveMeetingId = meetingId;
  _meetingDirty = false;
  _meetingStarted = false;
  _isNewMeeting = isNew;
  _manualSaveMode = manualSave;

  const container = document.getElementById('app');
  if (!container) return;

  if (manualSave) {
    let _dirtyCheckTimer: ReturnType<typeof setTimeout> | null = null;
    const checkAndShow = () => {
      if (_dirtyCheckTimer) clearTimeout(_dirtyCheckTimer);
      _dirtyCheckTimer = setTimeout(() => {
        const dirty = checkDirty();
        _meetingDirty = dirty;
        const btn = document.getElementById('btnSaveMeeting');
        if (btn) btn.style.display = dirty ? '' : 'none';
      }, 300);
    };
    container.addEventListener('input', checkAndShow, { signal });
    container.addEventListener('change', checkAndShow, { signal });
    container.addEventListener('click', checkAndShow, { signal });
    return;
  }

  const trigger = (delay = 3000) => {
    if (!_meetingStarted && _isNewMeeting) return;
    _meetingDirty = true;
    if (_autoSaveTimer) clearTimeout(_autoSaveTimer);
    _autoSaveTimer = setTimeout(() => doAutoSave(), delay);
  };

  container.addEventListener('input', () => trigger(3000), { signal });
  container.addEventListener('change', () => trigger(3000), { signal });
  container.addEventListener('click', () => trigger(5000), { signal });
}

/** Disable auto-save without flushing (used before deleting a meeting) */
export function disableAutoSave(): void {
  if (_autoSaveTimer) clearTimeout(_autoSaveTimer);
  _autoSaveDept = '';
  _autoSaveMeetingId = '';
  _meetingDirty = false;
  _meetingStarted = false;
  _isNewMeeting = false;
  _manualSaveMode = false;
}

export async function cleanupAutoSave(): Promise<void> {
  if (_autoSaveTimer) clearTimeout(_autoSaveTimer);
  // Auto-save on leave only if not in manual-save mode
  if (!_manualSaveMode && (_meetingDirty || _meetingStarted)) {
    await doAutoSave();
  }
  _autoSaveDept = '';
  _autoSaveMeetingId = '';
  _meetingDirty = false;
  _meetingStarted = false;
  _isNewMeeting = false;
  _manualSaveMode = false;
}

/** Force an immediate save */
export async function forceSave(): Promise<void> {
  if (_autoSaveTimer) clearTimeout(_autoSaveTimer);
  _meetingDirty = true;
  await doAutoSave();
  snapshotCleanState();
}

async function doAutoSave(): Promise<void> {
  if (!_autoSaveDept) return;
  if (!_meetingStarted && !_meetingDirty) return;
  try {
    const data = gatherMeetingData();
    if (_isNewMeeting && !_autoSaveMeetingId) {
      const result = await createMeeting(_autoSaveDept, data as Record<string, any>);
      if (!result) throw new Error('Create failed');
      _autoSaveMeetingId = result.id;
      _isNewMeeting = false;
      history.replaceState(null, '', `#/dept/${encodeURIComponent(_autoSaveDept)}/meeting/${_autoSaveMeetingId}`);
    } else {
      const ok = await saveMeeting(_autoSaveDept, _autoSaveMeetingId, data as Record<string, any>);
      if (!ok) throw new Error('Save failed');
    }
    showAutoSaved();
  } catch {
    /* silent */
  }
}

let _savedFadeTimer: ReturnType<typeof setTimeout> | null = null;

function showAutoSaved(): void {
  const el = document.getElementById('autosaveStatus');
  if (!el) return;
  if (_savedFadeTimer) clearTimeout(_savedFadeTimer);
  el.textContent = 'Saved';
  el.style.opacity = '1';
  _savedFadeTimer = setTimeout(() => { el.style.opacity = '0'; }, 3000);
}

// ── Scorecard / OKR Sync ──

let _syncTimer: ReturnType<typeof setTimeout> | null = null;

function syncScorecardToReview(): void {
  const fullRows = document.querySelectorAll('#scorecardFullTable tbody tr');
  const reviewRows = document.querySelectorAll('#scorecardTable tbody tr');
  fullRows.forEach((fullTr, i) => {
    if (i >= reviewRows.length) return;
    const fullEls = fullTr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    const revEls = reviewRows[i].querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    // scorecardFullTable: [name(0), owner(1), goal(2), wk1...wk13]
    // scorecardTable:     [name(0), owner(1), goal(2), actual(3), status(4), notes(5)]
    if (fullEls[0] && revEls[0] && revEls[0].value !== fullEls[0].value) revEls[0].value = fullEls[0].value;
    if (fullEls[1] && revEls[1] && revEls[1].value !== fullEls[1].value) revEls[1].value = fullEls[1].value;
  });
}

function syncReviewToScorecard(): void {
  const fullRows = document.querySelectorAll('#scorecardFullTable tbody tr');
  const reviewRows = document.querySelectorAll('#scorecardTable tbody tr');
  reviewRows.forEach((revTr, i) => {
    if (i >= fullRows.length) return;
    const revEls = revTr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    const fullEls = fullRows[i].querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    if (revEls[0] && fullEls[0] && fullEls[0].value !== revEls[0].value) fullEls[0].value = revEls[0].value;
    if (revEls[1] && fullEls[1] && fullEls[1].value !== revEls[1].value) fullEls[1].value = revEls[1].value;
  });
}

function syncOkrToReview(): void {
  const fullRows = document.querySelectorAll('#okrFullTable tbody tr');
  const reviewRows = document.querySelectorAll('#okrReviewTable tbody tr');
  fullRows.forEach((fullTr, i) => {
    if (i >= reviewRows.length) return;
    const fullEls = fullTr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    const revEls = reviewRows[i].querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    // okrFullTable:    [skip num td, desc(0), owner(1), due(2), priority(3), %done(4), status(5), notes(6)]
    // okrReviewTable:  [desc(0), owner(1), due(2), status(3), %done(4), notes(5)]
    if (fullEls[0] && revEls[0] && revEls[0].value !== fullEls[0].value) revEls[0].value = fullEls[0].value;
    if (fullEls[1] && revEls[1] && revEls[1].value !== fullEls[1].value) revEls[1].value = fullEls[1].value;
  });
}

function syncReviewToOkr(): void {
  const fullRows = document.querySelectorAll('#okrFullTable tbody tr');
  const reviewRows = document.querySelectorAll('#okrReviewTable tbody tr');
  reviewRows.forEach((revTr, i) => {
    if (i >= fullRows.length) return;
    const revEls = revTr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    const fullEls = fullRows[i].querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    if (revEls[0] && fullEls[0] && fullEls[0].value !== revEls[0].value) fullEls[0].value = revEls[0].value;
    if (revEls[1] && fullEls[1] && fullEls[1].value !== revEls[1].value) fullEls[1].value = revEls[1].value;
  });
}

export function setupScorecardOkrSync(): void {
  const debounceSync = (fn: () => void) => {
    if (_syncTimer) clearTimeout(_syncTimer);
    _syncTimer = setTimeout(fn, 100);
  };
  document.getElementById('scorecardFullTable')?.addEventListener('input', () => debounceSync(syncScorecardToReview));
  document.getElementById('scorecardTable')?.addEventListener('input', () => debounceSync(syncReviewToScorecard));
  document.getElementById('okrFullTable')?.addEventListener('input', () => debounceSync(syncOkrToReview));
  document.getElementById('okrReviewTable')?.addEventListener('input', () => debounceSync(syncReviewToOkr));
}

