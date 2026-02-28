import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { api } from "../client";
import type { WorkOrder } from "../types";

export function useWorkOrders(params?: { status?: string; priority?: string; archived?: boolean }) {
  const search = new URLSearchParams();
  if (params?.status) search.set("status", params.status);
  if (params?.priority) search.set("priority", params.priority);
  if (params?.archived !== undefined) search.set("archived", String(params.archived));
  const qs = search.toString();
  return useQuery({
    queryKey: ["work-orders", qs],
    queryFn: () => api.get<WorkOrder[]>(`/work-orders${qs ? `?${qs}` : ""}`),
  });
}

export function useCreateWorkOrder() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (body: Record<string, unknown>) => api.post<WorkOrder>("/work-orders", body),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["work-orders"] }),
  });
}

export function useUpdateWorkOrder(id: number) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (body: Record<string, unknown>) => api.put<WorkOrder>(`/work-orders/${id}`, body),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["work-orders"] }),
  });
}

export function useDeleteWorkOrder() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (id: number) => api.delete(`/work-orders/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["work-orders"] }),
  });
}

export function useArchiveWorkOrder() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (id: number) => api.post(`/work-orders/${id}/archive`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["work-orders"] }),
  });
}
