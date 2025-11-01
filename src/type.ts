export interface User {
  code: string;
  name: string;
  weightPct: number;
  online: boolean;
}

export interface TaskRow {
  [key: string]: unknown;
}

export interface AssignmentItem {
  userCode: string;
  userName: string;
  taskIndex: number;
}

export interface AllocationSummary {
  userCode: string;
  userName: string;
  weightPct: number;
  online: boolean;
  count: number;
}
