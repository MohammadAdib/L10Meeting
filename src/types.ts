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

export const DEFAULT_MEASURABLES = ['', '', '', '', '', '', ''];

/** Default row counts — matches blank.xlsx template slot counts */
export const DEFAULT_ROWS = {
  scorecard: 7,
  okr: 6,
  headlines: 7,
  todoReview: 7,
  issues: 16,
  idsIssues: 10,
  newTodos: 11,
  cascading: 6,
  rating: 10,
} as const;

/** Max rows per section — matches blank.xlsx template slot counts */
export const MAX_ROWS = {
  scorecardReview: 7,
  okrReview: 6,
  headlines: 7,
  todoReview: 7,
  issues: 16,
  idsBlocks: 10,
  idsTodos: 4,
  newTodos: 11,
  cascading: 6,
  rating: 10,
  scorecardFull: 10,
  okrFull: 10,
  keyResults: 3,
} as const;
