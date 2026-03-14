import { statusSelect, deleteBtn } from './utils';

let idsIssueCount = 0;

export function resetIdsIssueCount(): void {
  idsIssueCount = 0;
}

export function addScorecardRow(name = ''): void {
  const tb = document.querySelector('#scorecardTable tbody');
  if (!tb) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input value="${name}" placeholder="KPI name"></td><td><input placeholder="Owner"></td><td><input placeholder="Goal"></td><td><input placeholder="Actual"></td><td>${statusSelect(['', 'On Track', 'Off Track', 'At Risk'])}</td><td><input placeholder="Notes"></td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function addOkrReviewRow(name = ''): void {
  const tb = document.querySelector('#okrReviewTable tbody');
  if (!tb) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input value="${name}" placeholder="OKR description"></td><td><input placeholder="Owner"></td><td><input type="date"></td><td>${statusSelect(['', 'On Track', 'Off Track', 'At Risk'])}</td><td><input type="number" min="0" max="100" placeholder="%" style="width:60px"></td><td><input placeholder="Notes"></td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function addHeadlineRow(): void {
  const tb = document.querySelector('#headlinesTable tbody');
  if (!tb) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input placeholder="Headline"></td><td>${statusSelect(['', 'Customer', 'Employee'])}</td><td><input placeholder="Name"></td><td>${statusSelect(['', 'Yes', 'No'])}</td><td>${statusSelect(['', 'Yes', 'No'])}</td><td><input placeholder="Notes"></td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function addTodoReviewRow(): void {
  const tb = document.querySelector('#todoReviewTable tbody');
  if (!tb) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input placeholder="To-do item"></td><td><input placeholder="Owner"></td><td><input type="date"></td><td>${statusSelect(['', 'Done', 'Not Done'])}</td><td>${statusSelect(['', 'Yes', 'No'])}</td><td><input placeholder="Notes"></td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
  tr.querySelector('select')?.addEventListener('change', () => window.__updateTodoCompletion());
}

export function updateTodoCompletion(): void {
  const rows = document.querySelectorAll('#todoReviewTable tbody tr');
  let done = 0;
  rows.forEach(r => {
    const sel = r.querySelector('select') as HTMLSelectElement | null;
    if (sel?.value === 'Done') done++;
  });
  const el = document.getElementById('todoCompletionNum');
  if (el) el.textContent = `${done} / ${rows.length}`;
}

export function addIssueRow(): void {
  const tb = document.querySelector('#issuesListTable tbody');
  if (!tb) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input placeholder="Issue / obstacle"></td><td><input placeholder="Name"></td><td>${statusSelect(['', '1 - High', '2 - Medium', '3 - Low'])}</td><td>${statusSelect(['', 'New', 'In Progress', 'Resolved', 'Tabled'])}</td><td><input placeholder="e.g. 10 min" style="width:70px"></td><td>${statusSelect(['', 'Yes', 'No'])}</td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function addIDSIssue(): void {
  idsIssueCount++;
  const n = idsIssueCount;
  const container = document.getElementById('idsIssuesContainer');
  if (!container) return;
  const div = document.createElement('div');
  div.className = 'ids-issue';
  div.innerHTML = `
    <div class="ids-issue-header" onclick="this.nextElementSibling.classList.toggle('collapsed')">
      <h3><span class="ids-issue-num">${n}</span> Issue #${n}</h3>
      <button class="row-delete" onclick="event.stopPropagation();this.closest('.ids-issue').remove()">&times;</button>
    </div>
    <div>
      <div class="ids-field"><label>Issue</label><textarea placeholder="Describe the real issue (not the symptom)"></textarea></div>
      <div class="ids-field"><label>Root Cause</label><textarea placeholder="Ask 'why?' until you reach the root"></textarea></div>
      <div class="ids-field"><label>Solution</label><textarea placeholder="Agreed solution — be specific"></textarea></div>
      <label style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--text-muted);margin-top:8px;margin-bottom:12px;display:block;">New To-Do(s)</label>
      <table class="data-table" id="idsTodo-${n}">
        <thead><tr><th>To-Do</th><th>Owner</th><th>Due Date</th><th style="width:90px">Priority</th><th style="width:110px">Status</th><th>Notes</th><th style="width:30px"></th></tr></thead>
        <tbody></tbody>
      </table>
      <button class="btn btn-outline-dark btn-sm add-row-btn" onclick="window.__addIDSTodoRow(${n})">+ Add To-Do</button>
    </div>`;
  container.appendChild(div);
  for (let i = 0; i < 3; i++) addIDSTodoRow(n);
}

export function addIDSTodoRow(issueN: number): void {
  const tb = document.querySelector(`#idsTodo-${issueN} tbody`);
  if (!tb) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input placeholder="Action item"></td><td><input placeholder="Owner"></td><td><input type="date"></td><td>${statusSelect(['', 'High', 'Medium', 'Low'])}</td><td>${statusSelect(['', 'Not Started', 'In Progress', 'Done'])}</td><td><input placeholder="Notes"></td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function addNewTodoRow(): void {
  const tb = document.querySelector('#newTodoTable tbody');
  if (!tb) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input placeholder="Action item"></td><td><input placeholder="Owner"></td><td><input type="date"></td><td>${statusSelect(['', 'High', 'Medium', 'Low'])}</td><td>${statusSelect(['', 'Not Started', 'In Progress', 'Done'])}</td><td><input placeholder="Notes"></td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function addCascadingRow(): void {
  const tb = document.querySelector('#cascadingTable tbody');
  if (!tb) return;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td><input placeholder="Message"></td><td><input placeholder="To whom"></td><td><input type="date"></td><td><input placeholder="By whom"></td><td><input placeholder="e.g. Slack, Email"></td><td>${statusSelect(['', 'Yes', 'No'])}</td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function addRatingRow(): void {
  const tb = document.querySelector('#ratingTable tbody');
  if (!tb) return;
  const tr = document.createElement('tr');
  const stars = Array.from({ length: 10 }, (_, i) =>
    `<button onclick="window.__setRating(this,${i + 1})">&#9733;</button>`
  ).join('');
  tr.innerHTML = `<td><input placeholder="Name"></td><td><div class="rating-stars">${stars}</div><input type="hidden" class="rating-value" value="0"></td><td><input placeholder="Comment"></td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function setRating(btn: HTMLElement, val: number): void {
  const stars = btn.parentElement!.querySelectorAll('button');
  stars.forEach((s, i) => s.classList.toggle('active', i < val));
  const hidden = btn.parentElement!.nextElementSibling as HTMLInputElement;
  hidden.value = String(val);
  updateAvgRating();
}

export function updateAvgRating(): void {
  const inputs = document.querySelectorAll<HTMLInputElement>('#ratingTable .rating-value');
  let sum = 0, count = 0;
  inputs.forEach(inp => {
    const v = parseInt(inp.value);
    if (v > 0) { sum += v; count++; }
  });
  const el = document.getElementById('avgRating');
  if (el) el.textContent = count > 0 ? (sum / count).toFixed(1) : '\u2014';
}

// ── Scorecard Full Tab ──
export function addScorecardFullRow(name = ''): void {
  const tb = document.querySelector('#scorecardFullTable tbody');
  if (!tb) return;
  const tr = document.createElement('tr');
  let weeks = '';
  for (let i = 0; i < 13; i++) weeks += `<td><input placeholder="-" style="width:50px;text-align:center"></td>`;
  tr.innerHTML = `<td><input value="${name}" placeholder="KPI name"></td><td><input placeholder="Owner" style="width:80px"></td><td><input placeholder="Goal" style="width:60px"></td>${weeks}<td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

// ── OKR Full Tab ──
export function addOkrFullRow(name = '', num?: number): void {
  const tb = document.querySelector('#okrFullTable tbody');
  if (!tb) return;
  const n = num ?? tb.children.length + 1;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td style="width:30px;text-align:center;color:var(--text-muted)">${n}</td><td><input value="${name}" placeholder="OKR description"></td><td><input placeholder="Owner"></td><td><input type="date"></td><td>${statusSelect(['', 'High', 'Medium', 'Low'])}</td><td><input type="number" min="0" max="100" placeholder="%" style="width:55px"></td><td>${statusSelect(['', 'On Track', 'Off Track', 'At Risk'])}</td><td><input placeholder="Notes"></td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function addKeyResultRow(okrN: number, num?: number): void {
  const tb = document.querySelector(`#keyResults-${okrN} tbody`);
  if (!tb) return;
  const n = num ?? tb.children.length + 1;
  const tr = document.createElement('tr');
  tr.innerHTML = `<td style="width:30px;text-align:center;color:var(--text-muted)">${n}</td><td><input placeholder="Key result"></td><td><input placeholder="Owner"></td><td><input placeholder="Target"></td><td><input placeholder="Actual"></td><td><input type="number" min="0" max="100" placeholder="%" style="width:55px"></td><td>${statusSelect(['', 'On Track', 'Off Track', 'Done'])}</td><td><input placeholder="Notes"></td><td>${deleteBtn()}</td>`;
  tb.appendChild(tr);
}

export function buildKeyResultBlocks(): void {
  const container = document.getElementById('okrKeyResultsContainer');
  if (!container) return;
  for (let i = 1; i <= 3; i++) {
    const div = document.createElement('div');
    div.className = 'ids-issue';
    div.innerHTML = `
      <div class="ids-issue-header"><h3><span class="ids-issue-num">${i}</span> OKR #${i} — Key Results</h3></div>
      <table class="data-table" id="keyResults-${i}">
        <thead><tr><th>#</th><th>Key Result</th><th>Owner</th><th>Target</th><th>Actual</th><th style="width:70px">% Done</th><th style="width:110px">Status</th><th>Notes</th><th style="width:30px"></th></tr></thead>
        <tbody></tbody>
      </table>
      <button class="btn btn-outline-dark btn-sm add-row-btn" onclick="window.__addKeyResultRow(${i})">+ Add Key Result</button>`;
    container.appendChild(div);
    for (let j = 1; j <= 3; j++) addKeyResultRow(i, j);
  }
}
