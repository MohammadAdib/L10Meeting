/** Build a status <select> with color-coded options */
export function statusSelect(options: string[]): string {
  const opts = options.map(o => `<option value="${o}">${o}</option>`).join('');
  return `<select onchange="window.__onStatusChange(this)">${opts}</select>`;
}

/** Apply status color class to a <select> */
export function onStatusChange(sel: HTMLSelectElement): void {
  sel.className = '';
  const v = sel.value.toLowerCase();
  if (v === 'on track' || v === 'done' || v === 'yes' || v === 'solved') sel.className = 'status-done';
  else if (v === 'off track' || v === 'not done' || v === 'dropped') sel.className = 'status-off-track';
  else if (v === 'at risk' || v === 'carry over' || v === 'next meeting') sel.className = 'status-at-risk';
  else if (v === 'open') sel.className = 'status-at-risk';
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

/** Custom confirm dialog */
export function confirmDialog(message: string, confirmLabel = 'Confirm', destructive = false): Promise<boolean> {
  return new Promise(resolve => {
    const overlay = document.createElement('div');
    overlay.className = 'dialog-overlay';
    const dialog = document.createElement('div');
    dialog.className = 'dialog';
    dialog.innerHTML = `
      <p class="dialog-message">${message}</p>
      <div class="dialog-actions">
        <button class="btn btn-outline dialog-cancel">Cancel</button>
        <button class="btn ${destructive ? 'btn-danger' : 'btn-primary'} dialog-confirm">${confirmLabel}</button>
      </div>`;
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);
    requestAnimationFrame(() => overlay.classList.add('visible'));

    const close = (result: boolean) => {
      overlay.classList.remove('visible');
      setTimeout(() => overlay.remove(), 150);
      resolve(result);
    };

    overlay.addEventListener('click', (e) => { if (e.target === overlay) close(false); });
    dialog.querySelector('.dialog-cancel')!.addEventListener('click', () => close(false));
    dialog.querySelector('.dialog-confirm')!.addEventListener('click', () => close(true));
    (dialog.querySelector('.dialog-confirm') as HTMLElement).focus();
  });
}

/** Populate a table's rows from saved data arrays */
export function populateTableRows(tableSelector: string, rows: string[][]): void {
  const trs = document.querySelectorAll(`${tableSelector} tbody tr`);
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

/** Show toast notification */
export function showToast(msg: string): void {
  const t = document.querySelector<HTMLElement>('.toast');
  if (!t) return;
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2500);
}
