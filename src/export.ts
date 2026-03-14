import { showToast } from './utils';
import templateUrl from './template.xlsx?url';

/** Helper: get all input/select values from a table's tbody rows */
function getTableRows(tableId: string): string[][] {
  const rows: string[][] = [];
  document.querySelectorAll(`#${tableId} tbody tr`).forEach(tr => {
    const cells: string[] = [];
    tr.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea')
      .forEach(el => cells.push(el.value));
    rows.push(cells);
  });
  return rows;
}

function val(id: string): string {
  return (document.getElementById(id) as HTMLInputElement)?.value ?? '';
}

export async function exportExcel(): Promise<void> {
  showToast('Generating Excel...');

  const ExcelJS = await import('exceljs');
  const response = await fetch(templateUrl);
  const buffer = await response.arrayBuffer();

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const ws = wb.getWorksheet('L10 Meeting')!;

  // ── Meta ──
  ws.getCell('B2').value = val('metaTeam');
  ws.getCell('E2').value = val('metaDate');
  ws.getCell('B3').value = val('metaFacilitator');
  ws.getCell('E3').value = val('metaScribe');
  ws.getCell('B4').value = val('metaStart');
  ws.getCell('E4').value = val('metaEnd');

  // ── 1. Segue ──
  ws.getCell('B8').value = (document.getElementById('seguePersonal') as HTMLTextAreaElement)?.value ?? '';
  ws.getCell('B9').value = (document.getElementById('segueProfessional') as HTMLTextAreaElement)?.value ?? '';

  // ── 2. Scorecard Review (rows 14-20, cols A-F) ──
  const scorecard = getTableRows('scorecardTable');
  scorecard.forEach((row, i) => {
    if (i > 6) return;
    const r = 14 + i;
    // [name, owner, goal, actual, status, notes]
    ws.getCell(`A${r}`).value = row[0] || '';
    ws.getCell(`B${r}`).value = row[1] || '';
    ws.getCell(`C${r}`).value = row[2] || '';
    ws.getCell(`D${r}`).value = row[3] || '';
    ws.getCell(`E${r}`).value = row[4] || '';
    ws.getCell(`F${r}`).value = row[5] || '';
  });

  // ── 3. OKR Review (rows 26-31, cols A-F) ──
  const okrs = getTableRows('okrReviewTable');
  okrs.forEach((row, i) => {
    if (i > 5) return;
    const r = 26 + i;
    // [description, owner, due date, status, % done, notes]
    ws.getCell(`A${r}`).value = row[0] || '';
    ws.getCell(`B${r}`).value = row[1] || '';
    ws.getCell(`C${r}`).value = row[2] || '';
    ws.getCell(`D${r}`).value = row[3] || '';
    ws.getCell(`E${r}`).value = row[4] || '';
    ws.getCell(`F${r}`).value = row[5] || '';
  });

  // ── 4. Headlines (rows 37-42, cols A-F) ──
  const headlines = getTableRows('headlinesTable');
  headlines.forEach((row, i) => {
    if (i > 5) return;
    const r = 37 + i;
    // [headline, type, reported by, action needed, add to ids, notes]
    ws.getCell(`A${r}`).value = row[0] || '';
    ws.getCell(`B${r}`).value = row[1] || '';
    ws.getCell(`C${r}`).value = row[2] || '';
    ws.getCell(`D${r}`).value = row[3] || '';
    ws.getCell(`E${r}`).value = row[4] || '';
    ws.getCell(`F${r}`).value = row[5] || '';
  });

  // ── 5. To-Do Review (rows 47-53, cols A-F) ──
  const todos = getTableRows('todoReviewTable');
  todos.forEach((row, i) => {
    if (i > 6) return;
    const r = 47 + i;
    // [todo, owner, due date, status, add to ids, notes]
    ws.getCell(`A${r}`).value = row[0] || '';
    ws.getCell(`B${r}`).value = row[1] || '';
    ws.getCell(`C${r}`).value = row[2] || '';
    ws.getCell(`D${r}`).value = row[3] || '';
    ws.getCell(`E${r}`).value = row[4] || '';
    ws.getCell(`F${r}`).value = row[5] || '';
  });

  // Completion rate
  let done = 0;
  todos.forEach(r => { if (r[3] === 'Done') done++; });
  ws.getCell('E54').value = `${done} / ${todos.length} done`;

  // ── 6. IDS Issues List (rows 60-75, cols A-F) ──
  const issues = getTableRows('issuesListTable');
  issues.forEach((row, i) => {
    if (i > 15) return;
    const r = 60 + i;
    // [issue, raised by, priority, status, time est, next mtg]
    ws.getCell(`A${r}`).value = row[0] || '';
    ws.getCell(`B${r}`).value = row[1] || '';
    ws.getCell(`C${r}`).value = row[2] || '';
    ws.getCell(`D${r}`).value = row[3] || '';
    ws.getCell(`E${r}`).value = row[4] || '';
    ws.getCell(`F${r}`).value = row[5] || '';
  });

  // ── IDS Issue Detail Blocks ──
  const issueStarts = [77, 86, 95, 104, 113, 122, 131, 140, 149, 158];
  const idsBlocks = document.querySelectorAll('#idsIssuesContainer .ids-issue');
  idsBlocks.forEach((block, bi) => {
    if (bi >= issueStarts.length) return;
    const base = issueStarts[bi];
    const textareas = block.querySelectorAll<HTMLTextAreaElement>('.ids-field textarea');
    // issue, root cause, solution
    if (textareas[0]) ws.getCell(`B${base + 1}`).value = textareas[0].value;
    if (textareas[1]) ws.getCell(`B${base + 2}`).value = textareas[1].value;
    if (textareas[2]) ws.getCell(`B${base + 3}`).value = textareas[2].value;

    // To-dos for this issue (rows base+4 to base+8)
    const todoRows = getTableRows(`idsTodo-${bi + 1}`);
    todoRows.forEach((row, ti) => {
      if (ti > 4) return;
      const r = base + 4 + ti;
      ws.getCell(`A${r}`).value = row[0] || '';
      ws.getCell(`B${r}`).value = row[1] || '';
      ws.getCell(`C${r}`).value = row[2] || '';
      ws.getCell(`D${r}`).value = row[3] || '';
      ws.getCell(`E${r}`).value = row[4] || '';
      ws.getCell(`F${r}`).value = row[5] || '';
    });
  });

  // ── 7. Conclude — New To-Dos (rows 171-181, cols A-F) ──
  const newTodos = getTableRows('newTodoTable');
  newTodos.forEach((row, i) => {
    if (i > 10) return;
    const r = 171 + i;
    ws.getCell(`A${r}`).value = row[0] || '';
    ws.getCell(`B${r}`).value = row[1] || '';
    ws.getCell(`C${r}`).value = row[2] || '';
    ws.getCell(`D${r}`).value = row[3] || '';
    ws.getCell(`E${r}`).value = row[4] || '';
    ws.getCell(`F${r}`).value = row[5] || '';
  });

  // ── Cascading Messages (rows 184-189, cols A-F) ──
  const cascading = getTableRows('cascadingTable');
  cascading.forEach((row, i) => {
    if (i > 5) return;
    const r = 184 + i;
    ws.getCell(`A${r}`).value = row[0] || '';
    ws.getCell(`B${r}`).value = row[1] || '';
    ws.getCell(`C${r}`).value = row[2] || '';
    ws.getCell(`D${r}`).value = row[3] || '';
    ws.getCell(`E${r}`).value = row[4] || '';
    ws.getCell(`F${r}`).value = row[5] || '';
  });

  // ── Meeting Rating (rows 192-197, cols A-C) ──
  const ratingRows = document.querySelectorAll('#ratingTable tbody tr');
  let ratingSum = 0, ratingCount = 0;
  ratingRows.forEach((tr, i) => {
    if (i > 5) return;
    const r = 192 + i;
    const inputs = tr.querySelectorAll<HTMLInputElement>('input');
    const name = inputs[0]?.value || '';
    const ratingVal = tr.querySelector<HTMLInputElement>('.rating-value')?.value || '0';
    const comment = inputs[inputs.length - 1]?.value || '';
    ws.getCell(`A${r}`).value = name;
    ws.getCell(`B${r}`).value = parseInt(ratingVal) > 0 ? parseInt(ratingVal) : '';
    ws.getCell(`C${r}`).value = comment;
    const v = parseInt(ratingVal);
    if (v > 0) { ratingSum += v; ratingCount++; }
  });
  ws.getCell('B198').value = ratingCount > 0 ? (ratingSum / ratingCount).toFixed(1) : '';

  // ── Write & Download ──
  const outBuffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `L10_Meeting_${val('metaDate') || 'draft'}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);

  showToast('Excel exported!');
}
