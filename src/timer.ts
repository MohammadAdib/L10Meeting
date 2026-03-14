import type { TimerState } from './types';
import { SECTIONS } from './types';

export const timers: Record<number, TimerState> = {};

export function initTimers(): void {
  for (const sec of SECTIONS) {
    timers[sec.num] = { remaining: sec.time, running: false, interval: null };
    updateTimerDisplay(sec.num);
  }
}

export function toggleTimer(n: number): void {
  const t = timers[n];
  if (t.running) {
    if (t.interval) clearInterval(t.interval);
    t.running = false;
  } else {
    t.running = true;
    t.interval = setInterval(() => {
      t.remaining--;
      updateTimerDisplay(n);
      if (t.remaining <= 0 && t.remaining > -1) {
        // Keep running into negative to show overtime
      }
      updateProgress();
    }, 1000);
  }
  updateTimerBtn(n);
}

export function resetTimer(n: number): void {
  const t = timers[n];
  if (t.interval) clearInterval(t.interval);
  const sec = SECTIONS.find(s => s.num === n)!;
  timers[n] = { remaining: sec.time, running: false, interval: null };
  updateTimerDisplay(n);
  updateTimerBtn(n);
}

function updateTimerDisplay(n: number): void {
  const t = timers[n];
  const abs = Math.abs(t.remaining);
  const m = Math.floor(abs / 60);
  const s = abs % 60;
  const badge = document.getElementById(`timer-badge-${n}`);
  if (!badge) return;
  badge.textContent = (t.remaining < 0 ? '-' : '') + `${m}:${String(s).padStart(2, '0')}`;
  badge.className = 'timer-badge' + (t.running ? (t.remaining < 0 ? ' over' : ' running') : '');
}

function updateTimerBtn(n: number): void {
  const btn = document.getElementById(`timer-btn-${n}`);
  if (!btn) return;
  if (timers[n].running) {
    btn.innerHTML = '&#9646;&#9646;';
    btn.className = 'timer-btn timer-pause';
  } else {
    btn.innerHTML = '&#9654;';
    btn.className = 'timer-btn timer-play';
  }
}

export function updateProgress(): void {
  const total = SECTIONS.reduce((a, s) => a + s.time, 0);
  const elapsed = SECTIONS.reduce((sum, s) => sum + (s.time - timers[s.num].remaining), 0);
  const pct = Math.min(100, Math.max(0, (elapsed / total) * 100));
  const fill = document.getElementById('meetingProgress');
  if (fill) fill.style.width = pct + '%';
  const globalFill = document.getElementById('globalProgress');
  if (globalFill) globalFill.style.width = pct + '%';
}
