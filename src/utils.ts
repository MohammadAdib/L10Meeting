/** Build a status <select> with color-coded options */
export function statusSelect(options: string[]): string {
  const opts = options.map(o => `<option value="${o}">${o}</option>`).join('');
  return `<select onchange="window.__onStatusChange(this)">${opts}</select>`;
}

/** Apply status color class to a <select> */
export function onStatusChange(sel: HTMLSelectElement): void {
  sel.className = '';
  const v = sel.value.toLowerCase();
  if (v === 'on track' || v === 'done' || v === 'yes') sel.className = 'status-done';
  else if (v === 'off track' || v === 'not done') sel.className = 'status-off-track';
  else if (v === 'at risk') sel.className = 'status-at-risk';
}

/** Delete row button HTML */
export function deleteBtn(): string {
  return `<button class="row-delete" onclick="this.closest('tr').remove();window.__updateTodoCompletion();window.__updateAvgRating();">&times;</button>`;
}

/** Helper: get all input/select values from a table's tbody rows */
export function getTableRows(tableId: string): string[][] {
  const rows: string[][] = [];
  document.querySelectorAll(`#${tableId} tbody tr`).forEach(tr => {
    const cells: string[] = [];
    tr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea')
      .forEach(el => cells.push(el.value));
    rows.push(cells);
  });
  return rows;
}

/** Helper: get value of an input by id */
export function val(id: string): string {
  return (document.getElementById(id) as HTMLInputElement)?.value ?? '';
}

/** Show toast notification */
export function showToast(msg: string): void {
  const t = document.querySelector<HTMLElement>('.toast');
  if (!t) return;
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2500);
}
