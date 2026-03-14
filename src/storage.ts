import { showToast, getTableRows, val, populateTableRows } from './utils';
import {
  updateTodoCompletion, updateAvgRating,
  addScorecardRow, addOkrReviewRow, addHeadlineRow, addTodoReviewRow,
  addIssueRow, addIDSIssue, addIDSTodoRow, addNewTodoRow, addCascadingRow,
  addRatingRow, addScorecardFullRow, addOkrFullRow,
} from './tables';
import { createMeeting, saveMeeting, downloadMeetingExcel } from './fs-service';

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

export function setupAutoSave(dept: string, meetingId: string, isNew: boolean = false): void {
  _autoSaveDept = dept;
  _autoSaveMeetingId = meetingId;
  _meetingDirty = false;
  _meetingStarted = false;
  _isNewMeeting = isNew;

  const container = document.getElementById('app');
  if (!container) return;

  const trigger = (delay = 3000) => {
    if (!_meetingStarted && _isNewMeeting) return;
    _meetingDirty = true;
    if (_autoSaveTimer) clearTimeout(_autoSaveTimer);
    _autoSaveTimer = setTimeout(() => doAutoSave(), delay);
  };

  container.addEventListener('input', () => trigger(3000));
  container.addEventListener('change', () => trigger(3000));
  container.addEventListener('click', () => trigger(5000));
}

/** Disable auto-save without flushing (used before deleting a meeting) */
export function disableAutoSave(): void {
  if (_autoSaveTimer) clearTimeout(_autoSaveTimer);
  _autoSaveDept = '';
  _autoSaveMeetingId = '';
  _meetingDirty = false;
  _meetingStarted = false;
  _isNewMeeting = false;
}

export async function cleanupAutoSave(): Promise<void> {
  if (_autoSaveTimer) clearTimeout(_autoSaveTimer);
  // Save on leave if there are unsaved changes
  if (_meetingDirty || _meetingStarted) {
    await doAutoSave();
  }
  _autoSaveDept = '';
  _autoSaveMeetingId = '';
  _meetingDirty = false;
  _meetingStarted = false;
  _isNewMeeting = false;
}

/** Force an immediate save */
export async function forceSave(): Promise<void> {
  if (_autoSaveTimer) clearTimeout(_autoSaveTimer);
  _meetingDirty = true;
  await doAutoSave();
}

async function doAutoSave(): Promise<void> {
  if (!_autoSaveDept) return;
  if (!_meetingStarted && !_meetingDirty) return;
  try {
    const data = gatherMeetingData();
    data.lastSaved = new Date().toISOString();

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

/** Force a save then download the Excel file */
export async function openInExcel(): Promise<void> {
  await doAutoSave();
  if (!_autoSaveDept || !_autoSaveMeetingId) {
    showToast('Save the meeting first before downloading.');
    return;
  }
  try {
    await downloadMeetingExcel(_autoSaveDept, _autoSaveMeetingId);
    showToast('Downloading Excel file...');
  } catch {
    showToast('Could not download file.');
  }
}
