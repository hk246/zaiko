import { useQuery } from "@tanstack/react-query";
import { api } from "../client";
import type { CriticalPeriod, DashboardSummary, MaterialStats, TimelineEvent } from "../types";

export function useDashboardSummary() {
  return useQuery({
    queryKey: ["dashboard-summary"],
    queryFn: () => api.get<DashboardSummary>("/dashboard/summary"),
  });
}

export function useAlerts() {
  return useQuery({
    queryKey: ["dashboard-alerts"],
    queryFn: () => api.get<CriticalPeriod[]>("/dashboard/alerts"),
  });
}

export function useMaterialStats(materialId: number | string) {
  return useQuery({
    queryKey: ["material-stats", materialId],
    queryFn: () => api.get<MaterialStats[]>(`/dashboard/materials/${materialId}/stats`),
    enabled: !!materialId,
  });
}

export function useMaterialTimeline(materialId: number | string) {
  return useQuery({
    queryKey: ["material-timeline", materialId],
    queryFn: () => api.get<TimelineEvent[]>(`/dashboard/materials/${materialId}/timeline`),
    enabled: !!materialId,
  });
}
