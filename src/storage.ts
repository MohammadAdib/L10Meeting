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

/** Check if server API is available (desktop mode) */
export async function isServerAvailable(): Promise<boolean> {
  try {
    const res = await fetch('/api/meetings');
    return res.ok;
  } catch {
    return false;
  }
}

/** Save meeting to server */
export async function saveMeeting(): Promise<void> {
  const data = gatherMeetingData();
  const meta = data.meta as Record<string, string>;
  const date = meta.date || new Date().toISOString().split('T')[0];
  const team = (meta.team || 'Meeting').replace(/[^a-zA-Z0-9 ]/g, '').replace(/\s+/g, '_');
  const filename = `${date}_${team}.json`;

  try {
    const res = await fetch(`/api/meetings/${filename}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    });
    if (!res.ok) throw new Error('Save failed');
    showToast('Meeting saved!');
  } catch {
    showToast('Error saving meeting');
  }
}

/** List saved meetings from server */
export async function listMeetings(): Promise<{ filename: string; modified: string }[]> {
  try {
    const res = await fetch('/api/meetings');
    return await res.json();
  } catch {
    return [];
  }
}

/** Load a saved meeting from server */
export async function loadMeeting(filename: string): Promise<void> {
  try {
    const res = await fetch(`/api/meetings/${filename}`);
    if (!res.ok) throw new Error('Load failed');
    const data = await res.json();
    loadMeetingData(data);
    showToast('Meeting loaded!');
  } catch {
    showToast('Error loading meeting');
  }
}
