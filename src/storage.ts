import { showToast, getTableRows, val, onStatusChange } from './utils';
import { updateTodoCompletion, updateAvgRating } from './tables';

export function resetAll(): void {
  if (!confirm('Reset all fields? This cannot be undone.')) return;
  location.reload();
}

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
      quarter: val('okrQuarter'),
      year: val('okrYear'),
      startDate: val('okrStartDate'),
      targetDate: val('okrTargetDate'),
    },
  };

  // Gather all table data
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

  // Restore tables
  const tableIds = [
    'scorecardTable', 'okrReviewTable', 'headlinesTable', 'todoReviewTable',
    'issuesListTable', 'newTodoTable', 'cascadingTable', 'ratingTable',
    'scorecardFullTable', 'okrFullTable',
  ];
  for (const tableId of tableIds) {
    const rows = data[tableId] as string[][] | undefined;
    if (!rows) continue;
    const trs = document.querySelectorAll(`#${tableId} tbody tr`);
    rows.forEach((cells, ri) => {
      if (ri >= trs.length) return;
      const els = trs[ri].querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
      cells.forEach((v, ci) => {
        if (ci < els.length) {
          els[ci].value = v;
          if (els[ci] instanceof HTMLSelectElement) onStatusChange(els[ci] as HTMLSelectElement);
        }
      });
    });
  }

  // IDS blocks
  const idsBlocks = data.idsBlocks as { fields: string[]; todos: string[][] }[] | undefined;
  if (idsBlocks) {
    const blocks = document.querySelectorAll('#idsIssuesContainer .ids-issue');
    idsBlocks.forEach((block, bi) => {
      if (bi >= blocks.length) return;
      const tas = blocks[bi].querySelectorAll<HTMLTextAreaElement>('.ids-field textarea');
      block.fields.forEach((v, fi) => { if (fi < tas.length) tas[fi].value = v; });
    });
  }

  // Rating values
  const ratingValues = data.ratingValues as string[] | undefined;
  if (ratingValues) {
    const inputs = document.querySelectorAll<HTMLInputElement>('#ratingTable .rating-value');
    ratingValues.forEach((v, i) => { if (i < inputs.length) inputs[i].value = v; });
  }

  updateTodoCompletion();
  updateAvgRating();
}

/** Auto-save: debounced PUT to server */
let _autoSaveTimer: ReturnType<typeof setTimeout> | null = null;
let _autoSaveDept: string = '';
let _autoSaveMeetingId: string = '';

export function setupAutoSave(dept: string, meetingId: string): void {
  _autoSaveDept = dept;
  _autoSaveMeetingId = meetingId;

  // Listen for any input changes within the meeting view
  const container = document.getElementById('app');
  if (!container) return;

  const trigger = () => {
    if (_autoSaveTimer) clearTimeout(_autoSaveTimer);
    updateAutoSaveStatus('Unsaved changes...');
    _autoSaveTimer = setTimeout(() => doAutoSave(), 3000);
  };

  container.addEventListener('input', trigger);
  container.addEventListener('change', trigger);
}

async function doAutoSave(): Promise<void> {
  if (!_autoSaveDept || !_autoSaveMeetingId) return;
  updateAutoSaveStatus('Saving...');
  try {
    const data = gatherMeetingData();
    data.lastSaved = new Date().toISOString();
    const res = await fetch(`/api/departments/${encodeURIComponent(_autoSaveDept)}/meetings/${_autoSaveMeetingId}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    });
    if (!res.ok) throw new Error('Save failed');
    updateAutoSaveStatus('Saved');
  } catch {
    updateAutoSaveStatus('Save failed');
  }
}

function updateAutoSaveStatus(text: string): void {
  const el = document.getElementById('autosaveStatus');
  if (el) el.textContent = text;
}
