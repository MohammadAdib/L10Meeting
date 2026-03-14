export interface TimerState {
  remaining: number;
  running: boolean;
  interval: ReturnType<typeof setInterval> | null;
}

export interface SectionConfig {
  num: number;
  name: string;
  time: number;
}

export const SECTIONS: SectionConfig[] = [
  { num: 1, name: 'Segue', time: 300 },
  { num: 2, name: 'Scorecard', time: 300 },
  { num: 3, name: 'OKR Review', time: 300 },
  { num: 4, name: 'Headlines', time: 300 },
  { num: 5, name: 'To-Do Review', time: 300 },
  { num: 6, name: 'IDS', time: 3600 },
  { num: 7, name: 'Conclude', time: 300 },
];

export const DEFAULT_MEASURABLES = ['', '', ''];
