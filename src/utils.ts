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
  else if (v === 'off track' || v === 'not done' || v === 'dropped' || v === "won't do") sel.className = 'status-off-track';
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
        // Update person picker button text if this is a hidden person-value input
        if (els[ci].classList.contains('person-value')) {
          const picker = els[ci].parentElement;
          const btn = picker?.querySelector('.person-picker-btn');
          if (btn) btn.textContent = v || '';
          if (picker) picker.classList.toggle('has-value', !!v);
        }
      }
    });
  });
}

// ── Person Picker (multi-select with checkboxes) ──

let _activePickerDropdown: HTMLElement | null = null;

function closeActivePicker(): void {
  if (_activePickerDropdown) {
    _activePickerDropdown.remove();
    _activePickerDropdown = null;
  }
}

let _pickerAbort: AbortController | null = null;

export function initPersonPickers(): void {
  // Abort previous listeners to avoid duplicates across navigations
  if (_pickerAbort) _pickerAbort.abort();
  _pickerAbort = new AbortController();
  const { signal } = _pickerAbort;

  // Clear button handler
  document.addEventListener('click', (e) => {
    const clearBtn = (e.target as HTMLElement).closest('.person-picker-clear') as HTMLElement | null;
    if (clearBtn) {
      e.stopPropagation();
      const picker = clearBtn.closest('.person-picker') as HTMLElement;
      if (!picker) return;
      const input = picker.querySelector<HTMLInputElement>('.person-value');
      const btn = picker.querySelector<HTMLElement>('.person-picker-btn');
      if (input) {
        input.value = '';
        input.dispatchEvent(new Event('change', { bubbles: true }));
      }
      if (btn) btn.textContent = '';
      picker.classList.remove('has-value');
      closeActivePicker();
      return;
    }
  }, { signal });

  document.addEventListener('click', (e) => {
    const btn = (e.target as HTMLElement).closest('.person-picker-btn') as HTMLElement | null;
    if (btn) {
      e.stopPropagation();
      const picker = btn.closest('.person-picker') as HTMLElement;
      if (!picker || picker.classList.contains('disabled')) return;

      // If this picker's dropdown is already open, close it
      if (_activePickerDropdown && _activePickerDropdown.dataset.pickerId === getPickerId(picker)) {
        closeActivePicker();
        return;
      }
      closeActivePicker();
      openPickerDropdown(picker, btn);
      return;
    }

    // Close if click is outside dropdown
    if (_activePickerDropdown && !(e.target as HTMLElement).closest('.person-picker-dropdown')) {
      closeActivePicker();
    }
  }, { signal });
}

let _pickerIdCounter = 0;
function getPickerId(picker: HTMLElement): string {
  if (!picker.dataset.pickerId) picker.dataset.pickerId = String(++_pickerIdCounter);
  return picker.dataset.pickerId;
}

function openPickerDropdown(picker: HTMLElement, btn: HTMLElement): void {
  const people = (picker.dataset.people || '').split('|||').filter(Boolean);
  const hiddenInput = picker.querySelector<HTMLInputElement>('.person-value')!;
  const isAllValue = hiddenInput.value.trim() === 'All';
  const currentValues = isAllValue ? [...people] : (hiddenInput.value ? hiddenInput.value.split(', ').map(s => s.trim()).filter(Boolean) : []);

  const dropdown = document.createElement('div');
  dropdown.className = 'person-picker-dropdown';
  dropdown.dataset.pickerId = getPickerId(picker);

  // Custom input row
  const customRow = document.createElement('div');
  customRow.className = 'pp-custom-row';
  customRow.innerHTML = `<input type="text" class="pp-custom-input" placeholder="Type a name..."><button type="button" class="pp-custom-add">+</button>`;
  dropdown.appendChild(customRow);

  if (people.length > 0) {
    const divider = document.createElement('div');
    divider.className = 'pp-divider';
    dropdown.appendChild(divider);
  }

  // People checkboxes
  const listEl = document.createElement('div');
  listEl.className = 'pp-list';
  for (const p of people) {
    const checked = currentValues.includes(p);
    const item = document.createElement('label');
    item.className = 'pp-item';
    item.innerHTML = `<input type="checkbox" ${checked ? 'checked' : ''}><span>${p}</span>`;
    listEl.appendChild(item);
  }
  dropdown.appendChild(listEl);

  // "All" option at bottom
  if (people.length > 1) {
    const allDivider = document.createElement('div');
    allDivider.className = 'pp-divider';
    dropdown.appendChild(allDivider);

    const allItem = document.createElement('label');
    allItem.className = 'pp-item pp-all';
    const allChecked = people.length > 0 && people.every(p => currentValues.includes(p));
    allItem.innerHTML = `<input type="checkbox" class="pp-all-cb" ${allChecked ? 'checked' : ''}><span>All</span>`;
    dropdown.appendChild(allItem);
  }

  // Position dropdown below the button
  const rect = btn.getBoundingClientRect();
  dropdown.style.position = 'fixed';
  dropdown.style.top = `${rect.bottom + 2}px`;
  dropdown.style.left = `${rect.left}px`;
  document.body.appendChild(dropdown);
  _activePickerDropdown = dropdown;

  // Adjust if going off-screen
  const dropRect = dropdown.getBoundingClientRect();
  if (dropRect.right > window.innerWidth) {
    dropdown.style.left = `${window.innerWidth - dropRect.width - 8}px`;
  }
  if (dropRect.bottom > window.innerHeight) {
    dropdown.style.top = `${rect.top - dropRect.height - 2}px`;
  }

  // Dismiss on scroll (ignore scrolls inside the dropdown itself)
  const onScroll = (e: Event) => {
    if (dropdown.contains(e.target as Node)) return;
    closeActivePicker();
    document.removeEventListener('scroll', onScroll, true);
  };
  setTimeout(() => {
    document.addEventListener('scroll', onScroll, true);
  }, 200);

  function updateValue(): void {
    const checks = listEl.querySelectorAll<HTMLInputElement>('input[type="checkbox"]');
    const selected: string[] = [];
    checks.forEach((cb, i) => { if (cb.checked) selected.push(people[i]); });
    // Also include any custom entries that were in the value but not in people list
    const customEntries = currentValues.filter(v => !people.includes(v));
    const allSelected = [...selected, ...customEntries];
    const isAll = people.length > 0 && people.every(p => selected.includes(p)) && customEntries.length === 0;
    hiddenInput.value = isAll ? 'All' : allSelected.join(', ');
    btn.textContent = isAll ? 'All' : (allSelected.length > 0 ? allSelected.join(', ') : '');
    picker.classList.toggle('has-value', allSelected.length > 0);
    hiddenInput.dispatchEvent(new Event('change', { bubbles: true }));

    // Update "All" checkbox
    const allCb = dropdown.querySelector<HTMLInputElement>('.pp-all-cb');
    if (allCb) allCb.checked = people.every(p => selected.includes(p));
  }

  // Checkbox change handlers
  listEl.addEventListener('change', () => updateValue());

  // "All" toggle
  const allCb = dropdown.querySelector<HTMLInputElement>('.pp-all-cb');
  if (allCb) {
    allCb.addEventListener('change', () => {
      const checks = listEl.querySelectorAll<HTMLInputElement>('input[type="checkbox"]');
      checks.forEach(cb => cb.checked = allCb.checked);
      updateValue();
    });
  }

  // Custom name add
  const customInput = dropdown.querySelector<HTMLInputElement>('.pp-custom-input')!;
  const customAddBtn = dropdown.querySelector<HTMLButtonElement>('.pp-custom-add')!;

  function addCustom(): void {
    const name = customInput.value.trim();
    if (!name) return;
    // Add to current values if not already there
    const cur = hiddenInput.value ? hiddenInput.value.split(', ').filter(Boolean) : [];
    if (!cur.includes(name)) {
      cur.push(name);
      hiddenInput.value = cur.join(', ');
      btn.textContent = cur.join(', ');
      picker.classList.add('has-value');
      currentValues.push(name);
      hiddenInput.dispatchEvent(new Event('change', { bubbles: true }));
    }
    customInput.value = '';
    // Check the checkbox if this person is in the list
    const idx = people.indexOf(name);
    if (idx >= 0) {
      const checks = listEl.querySelectorAll<HTMLInputElement>('input[type="checkbox"]');
      if (checks[idx]) checks[idx].checked = true;
    }
  }

  customAddBtn.addEventListener('click', (e) => { e.stopPropagation(); addCustom(); });
  customInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addCustom(); } });
  customInput.focus();
}

/** Show toast notification */
export function showToast(msg: string): void {
  const t = document.querySelector<HTMLElement>('.toast');
  if (!t) return;
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2500);
}
