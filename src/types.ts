export type Status = 'ok' | 'error';

export interface MetricHistory {
  timestamp: string;
  value: number;
}

export interface Metric {
  id: string;
  title: string;
  value: number | string;
  status: Status;
  lastUpdate: string;
  lastUpdateAt?: string; // ISO string for precise calculation
  refreshInterval?: number; // in minutes
  objective?: string;
  rules?: string[];
  history?: number[];
  details?: any[];
  isDynamic?: boolean;
  sqlQuery?: string;
  cachedData?: string;
  orderIndex?: number;
}

export interface Section {
  id?: string;
  title: string;
  metrics: Metric[];
  orderIndex?: number;
}
