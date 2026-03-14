import { showToast } from './utils';
import { updateTodoCompletion, updateAvgRating } from './tables';
import { onStatusChange } from './utils';

const STORAGE_KEY = 'l10-draft';

export function gatherData(): Record<string, unknown> {
  const getValue = (id: string) => (document.getElementById(id) as HTMLInputElement)?.value ?? '';

  const data: Record<string, unknown> = {
    meta: {
      team: getValue('metaTeam'),
      date: getValue('metaDate'),
      facilitator: getValue('metaFacilitator'),
      scribe: getValue('metaScribe'),
      start: getValue('metaStart'),
      end: getValue('metaEnd'),
    },
    segue: {
      personal: (document.getElementById('seguePersonal') as HTMLTextAreaElement)?.value ?? '',
      professional: (document.getElementById('segueProfessional') as HTMLTextAreaElement)?.value ?? '',
    },
  };

  // Gather all table data
  document.querySelectorAll<HTMLTableElement>('.data-table').forEach(table => {
    if (!table.id) return;
    const rows: string[][] = [];
    table.querySelectorAll('tbody tr').forEach(tr => {
      const cells: string[] = [];
      tr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea')
        .forEach(el => cells.push(el.value));
      rows.push(cells);
    });
    data[table.id] = rows;
  });

  // IDS issue text fields
  const idsBlocks: string[][] = [];
  document.querySelectorAll('#idsIssuesContainer .ids-issue').forEach(block => {
    const fields: string[] = [];
    block.querySelectorAll<HTMLTextAreaElement>('.ids-field textarea').forEach(ta => fields.push(ta.value));
    idsBlocks.push(fields);
  });
  data.idsBlocks = idsBlocks;

  return data;
}

export function saveDraft(): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(gatherData()));
  showToast('Draft saved!');
}

export function loadDraft(): void {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) { showToast('No draft found'); return; }

  const data = JSON.parse(raw) as Record<string, unknown>;

  // Meta
  const meta = data.meta as Record<string, string> | undefined;
  if (meta) {
    (['team', 'date', 'facilitator', 'scribe', 'start', 'end'] as const).forEach(key => {
      const idMap: Record<string, string> = { team: 'metaTeam', date: 'metaDate', facilitator: 'metaFacilitator', scribe: 'metaScribe', start: 'metaStart', end: 'metaEnd' };
      const el = document.getElementById(idMap[key]) as HTMLInputElement | null;
      if (el) el.value = meta[key] ?? '';
    });
  }

  // Segue
  const segue = data.segue as Record<string, string> | undefined;
  if (segue) {
    const p = document.getElementById('seguePersonal') as HTMLTextAreaElement | null;
    const pr = document.getElementById('segueProfessional') as HTMLTextAreaElement | null;
    if (p) p.value = segue.personal ?? '';
    if (pr) pr.value = segue.professional ?? '';
  }

  // Tables
  for (const [tableId, rows] of Object.entries(data)) {
    if (tableId === 'meta' || tableId === 'segue' || tableId === 'idsBlocks') continue;
    const table = document.getElementById(tableId);
    if (!table) continue;
    const trs = table.querySelectorAll('tbody tr');
    (rows as string[][]).forEach((cells, ri) => {
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
  const idsBlocks = data.idsBlocks as string[][] | undefined;
  if (idsBlocks) {
    const blocks = document.querySelectorAll('#idsIssuesContainer .ids-issue');
    idsBlocks.forEach((fields, bi) => {
      if (bi >= blocks.length) return;
      const tas = blocks[bi].querySelectorAll<HTMLTextAreaElement>('.ids-field textarea');
      fields.forEach((v, fi) => { if (fi < tas.length) tas[fi].value = v; });
    });
  }

  updateTodoCompletion();
  updateAvgRating();
  showToast('Draft loaded!');
}

export function resetAll(): void {
  if (!confirm('Reset all fields? This cannot be undone.')) return;
  localStorage.removeItem(STORAGE_KEY);
  location.reload();
}
