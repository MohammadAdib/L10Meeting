import { SECTIONS } from './types';
import { getLogoUrl } from './logo';

function sectionCard(num: number, title: string, timeLabel: string, bodyId: string, bodyHTML: string): string {
  const sec = SECTIONS.find(s => s.num === num);
  const mins = sec ? Math.round(sec.time / 60) : 0;
  return `
  <div class="section-card" id="sec-${num}">
    <div class="section-header" data-section="${num}">
      <h2><span class="section-num">${num}</span> ${title} <span class="chevron open" id="chev-${num}">&#9662;</span></h2>
      <span class="section-duration" id="duration-${num}">&#9201; ${mins} min</span>
      <div class="section-timer">
        <span class="timer-badge" id="timer-badge-${num}">${timeLabel}</span>
        <button class="timer-btn timer-play" id="timer-btn-${num}" data-timer="${num}">&#9654;</button>
        <button class="timer-btn timer-reset" data-timer-reset="${num}">&#8634;</button>
      </div>
    </div>
    <div class="section-body-wrap" id="${bodyId}"><div class="section-body">${bodyHTML}</div></div>
  </div>`;
}

export function tableHTML(id: string, headers: string[]): string {
  const ths = headers.map(h => {
    if (h.startsWith('w:')) {
      const [, w, label] = h.match(/w:(\d+):(.*)/)!;
      return `<th style="width:${w}px">${label}</th>`;
    }
    return `<th>${h}</th>`;
  }).join('');
  return `<table class="data-table" id="${id}"><thead><tr>${ths}</tr></thead><tbody></tbody></table>`;
}

/** Reusable Scorecard section HTML (used in both meeting tab and dept view) */
export function buildScorecardContent(): string {
  return `
    <div class="section-card">
      <div class="section-header" style="cursor:default"><h2>SCORECARD TRACKER (Rolling 13 Weeks)</h2></div>
      <div class="section-body">
        <p class="section-desc">Track weekly actuals below. Use color-coded status to flag off-track items.</p>
        ${tableHTML('scorecardFullTable', ['Measurable / KPI', 'Owner', 'Goal', 'Wk 1', 'Wk 2', 'Wk 3', 'Wk 4', 'Wk 5', 'Wk 6', 'Wk 7', 'Wk 8', 'Wk 9', 'Wk 10', 'Wk 11', 'Wk 12', 'Wk 13', 'w:30:'])}
        <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddScorecardFull">+ Add Measurable</button>
      </div>
    </div>`;
}

/** Reusable OKR section HTML (used in both meeting tab and dept view) */
export function buildOkrsContent(): string {
  return `
    <div class="section-card">
      <div class="section-header" style="cursor:default"><h2>OKR TRACKER (Rocks / 90-Day Priorities)</h2></div>
      <div class="section-body">
        <div class="meta-grid" style="grid-template-columns:1fr 1fr 1fr 1fr;margin-top:16px;margin-bottom:16px;">
          <div class="meta-field"><label>Quarter</label><select id="okrQuarter">${[1,2,3,4].map(q => `<option${q === Math.ceil((new Date().getMonth()+1)/3) ? ' selected' : ''}>Q${q}</option>`).join('')}</select></div>
          <div class="meta-field"><label>Year</label><select id="okrYear">${Array.from({length: 7}, (_, i) => { const y = new Date().getFullYear() - 2 + i; return `<option${y === new Date().getFullYear() ? ' selected' : ''}>${y}</option>`; }).join('')}</select></div>
          <div class="meta-field"><label>Start Date</label><input id="okrStartDate" type="date"></div>
          <div class="meta-field"><label>Target Completion</label><input id="okrTargetDate" type="date"></div>
        </div>
        <p class="section-desc">Each owner reports On Track / Off Track weekly in the L10. Off-track items go to IDS.</p>
        <div style="overflow-x:auto">${tableHTML('okrFullTable', ['#', 'OKR / Rock Description', 'Owner', 'Due Date', 'w:90:Priority', 'w:70:% Done', 'w:110:Status', 'Notes', 'w:30:'])}</div>
        <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddOkrFull">+ Add OKR</button>
        <div id="okrKeyResultsContainer" style="margin-top:24px;"></div>
      </div>
    </div>`;
}

export function buildAppHTML(deptName?: string, standalone = false): string {
  const sidebarItems = SECTIONS.map((s, i) =>
    `<a class="sidebar-item${i === 0 ? ' active' : ''}" data-nav="${s.num}" href="#sec-${s.num}">
      <span class="sidebar-num">${s.num}</span>
      <span class="sidebar-label">${s.name}</span>
    </a>`
  ).join('');

  return `
<div class="top-bar-wrapper">
  <div class="top-bar"${standalone ? ' style="padding-left:0"' : ''}>
    <div class="top-bar-left">
      ${standalone
        ? ''
        : (getLogoUrl() ? `<img src="${getLogoUrl()}" class="top-bar-logo">` : `<button class="top-bar-logo-placeholder" id="btnAddLogo">+ Add Logo</button>`)
      }
      <div class="top-bar-tabs">
        <button class="top-tab active" data-tab="meeting">L10 Meeting</button>
        <button class="top-tab" data-tab="scorecard">Scorecard</button>
        <button class="top-tab" data-tab="okrs">OKRs</button>
      </div>
    </div>
    <div class="top-bar-actions" id="topBarActions" style="opacity:0;pointer-events:none">
      <span class="autosave-status" id="autosaveStatus"></span>
      ${standalone ? `
        <button class="btn-export" id="btnExportExcel">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/></svg>
          Export
        </button>
        <button class="settings-btn" id="btnStandaloneBack" title="Back">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>
      ` : `
        <button class="btn btn-danger" id="btnDeleteMeeting" style="display:none">Delete</button>
      `}
    </div>
  </div>
  <div class="global-progress"><div class="global-progress-fill" id="globalProgress"></div></div>
</div>

<div class="app-layout">
  <!-- LEFT SIDEBAR -->
  <nav class="sidebar" id="sidebar">
    <div class="meeting-control" id="meetingControl">
      <button class="meeting-start-btn" id="btnMeetingStart">&#9654; Start Meeting</button>
    </div>
    ${sidebarItems}
  </nav>

  <div class="main-content">
    <div class="container">

      <!-- MEETING TAB -->
      <div id="tab-meeting" class="tab-content active">

    <div class="meta-grid">
      <div class="meta-field"><label>Team</label><input id="metaTeam" placeholder="e.g. Leadership Team"></div>
      <div class="meta-field"><label>Date</label><input id="metaDate" type="date"></div>
      <div class="meta-field"><label>Facilitator</label><input id="metaFacilitator" placeholder="Name"></div>
      <div class="meta-field"><label>Scribe</label><input id="metaScribe" placeholder="Name"></div>
      <div class="meta-field"><label>Start Time</label><input id="metaStart" type="time"></div>
      <div class="meta-field"><label>End Time</label><input id="metaEnd" type="time"></div>
    </div>

    ${sectionCard(1, 'SEGUE', '5:00', 'body-1', `
      <p class="section-desc">Each person shares one personal win and one professional win to open the meeting.</p>
      <div class="segue-grid">
        <div class="segue-box"><h3>Personal Good News</h3><textarea id="seguePersonal" placeholder="Share personal wins..."></textarea></div>
        <div class="segue-box"><h3>Professional Good News</h3><textarea id="segueProfessional" placeholder="Share professional wins..."></textarea></div>
      </div>
    `)}

    ${sectionCard(2, 'SCORECARD REVIEW', '5:00', 'body-2', `
      <p class="section-desc">Review each measurable. Off-track items → add to IDS.</p>
      ${tableHTML('scorecardTable', ['w:300:Measurable / KPI', 'Owner', 'Goal', 'Actual', 'w:110:Status', 'Notes', 'w:30:'])}
      <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddScorecard">+ Add Measurable</button>
    `)}

    ${sectionCard(3, 'OKR REVIEW', '5:00', 'body-3', `
      <p class="section-desc">Report On Track / Off Track for each OKR. Off-track items → add to IDS.</p>
      ${tableHTML('okrReviewTable', ['OKR / Rock Description', 'Owner', 'Due Date', 'w:110:Status', 'w:70:% Done', 'Notes', 'w:30:'])}
      <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddOkrReview">+ Add OKR</button>
    `)}

    ${sectionCard(4, 'CUSTOMER / EMPLOYEE HEADLINES', '5:00', 'body-4', `
      <p class="section-desc">Share good or bad news about customers or employees. Drop issues into IDS.</p>
      ${tableHTML('headlinesTable', ['Headline', 'w:120:Type', 'w:100:Reported By', 'w:110:Action Needed?', 'w:100:Add to IDS?', 'Notes', 'w:30:'])}
      <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddHeadline">+ Add Headline</button>
    `)}

    ${sectionCard(5, 'TO-DO LIST REVIEW', '5:00', 'body-5', `
      <p class="section-desc">Review last week's 7-day commitments. Done / Not Done. Incomplete items → IDS.</p>
      ${tableHTML('todoReviewTable', ["Last Week's To-Do", 'Owner', 'Due Date', 'w:110:Status', 'w:100:Add to IDS?', 'Notes', 'w:30:'])}
      <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddTodoReview">+ Add To-Do</button>
      <div class="completion-stat">
        <span class="stat-num" id="todoCompletionNum">0 / 0</span>
        <span class="stat-label">Completion Rate</span>
      </div>
    `)}

    ${sectionCard(6, 'IDS — IDENTIFY, DISCUSS, SOLVE', '60:00', 'body-6', `
      <p class="section-desc">IDENTIFY — Build the Issues List (vote to prioritize top 3 before discussing)</p>
      ${tableHTML('issuesListTable', ['w:250:Issue / Obstacle', 'w:130:Raised By', 'w:90:Priority', 'w:110:Status', 'w:80:Time Est.', 'w:90:Next Mtg?', 'w:30:'])}
      <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddIssue">+ Add Issue</button>
      <h3 style="margin:24px 0 12px;font-size:14px;color:var(--text-dim);">DISCUSS & SOLVE — IDS each issue completely before moving to the next</h3>
      <div id="idsIssuesContainer"></div>
      <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddIDSIssue">+ Add Issue Detail Block</button>
    `)}

    ${sectionCard(7, 'CONCLUDE', '5:00', 'body-7', `
      <h3 class="sub-heading" style="margin-top:14px;">New To-Do List — This Week's Commitments</h3>
      ${tableHTML('newTodoTable', ['To-Do', 'Owner', 'Due Date', 'w:90:Priority', 'w:110:Status', 'Notes', 'w:30:'])}
      <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddNewTodo">+ Add To-Do</button>

      <h3 class="sub-heading sub-heading-spaced">Cascading Messages — What needs to be shared?</h3>
      ${tableHTML('cascadingTable', ['Message', 'To Whom', 'By When', 'By Whom', 'Channel', 'w:80:Done?', 'w:30:'])}
      <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddCascading">+ Add Message</button>

      <h3 class="sub-heading sub-heading-spaced">Meeting Rating — Rate 1-10 (discuss anything below 8)</h3>
      ${tableHTML('ratingTable', ['Team Member', 'w:200:Rating (1-10)', 'Quick Comment', 'w:30:'])}
      <button class="btn btn-outline-dark btn-sm add-row-btn" id="btnAddRating">+ Add Member</button>
      <div class="completion-stat" style="margin-top:12px;">
        <span class="stat-num" id="avgRating">—</span>
        <span class="stat-label">Average Rating</span>
      </div>
    `)}
  </div>

  <!-- SCORECARD TAB -->
  <div id="tab-scorecard" class="tab-content">
    ${buildScorecardContent()}
  </div>

  <!-- OKR TAB -->
  <div id="tab-okrs" class="tab-content">
    ${buildOkrsContent()}
    </div>
  </div>
  </div><!-- end main-content -->
</div><!-- end app-layout -->

<div class="toast"></div>`;
}
