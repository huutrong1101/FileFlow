export interface User {
  code: string;
  name: string;
  weightPct: number;
  online: boolean;
  warehouses?: string[];
  order?: number;
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

export type MonthAgg = {
  expectedCum: Record<string, number>;
  actualCum: Record<string, number>;
  deficit: Record<string, number>;
  lastServedAt: Record<string, string>;
  version?: string;
};
