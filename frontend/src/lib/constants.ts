export const API_BASE = "/api";

export const STATUS_LABELS: Record<string, string> = {
  pending: "未実行",
  executed: "実行済",
  cancelled: "キャンセル",
  in_progress: "進行中",
  completed: "完了",
};

export const PRIORITY_LABELS: Record<string, string> = {
  critical: "緊急",
  high: "高",
  low: "低",
  none: "なし",
};

export const PRIORITY_COLORS: Record<string, string> = {
  critical: "bg-red-100 text-red-800",
  high: "bg-orange-100 text-orange-800",
  low: "bg-blue-100 text-blue-800",
  none: "bg-gray-100 text-gray-600",
};

export const STATUS_COLORS: Record<string, string> = {
  pending: "bg-yellow-100 text-yellow-800",
  in_progress: "bg-blue-100 text-blue-800",
  completed: "bg-green-100 text-green-800",
  executed: "bg-green-100 text-green-800",
  cancelled: "bg-gray-100 text-gray-600",
};

export const STAT_PERIODS = [1, 7, 30, 90, 180, 365] as const;
