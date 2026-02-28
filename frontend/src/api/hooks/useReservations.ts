import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { api } from "../client";
import type { Reservation } from "../types";

export function useReservations(params?: { status?: string; type?: string; material_id?: number }) {
  const search = new URLSearchParams();
  if (params?.status) search.set("status", params.status);
  if (params?.type) search.set("type", params.type);
  if (params?.material_id) search.set("material_id", String(params.material_id));
  const qs = search.toString();
  return useQuery({
    queryKey: ["reservations", qs],
    queryFn: () => api.get<Reservation[]>(`/reservations${qs ? `?${qs}` : ""}`),
  });
}

export function useCreateReservation() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (body: Record<string, unknown>) => api.post<Reservation>("/reservations", body),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["reservations"] });
      qc.invalidateQueries({ queryKey: ["materials"] });
    },
  });
}

export function useExecuteReservation() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ id, body }: { id: number; body: Record<string, unknown> }) =>
      api.post<Reservation>(`/reservations/${id}/execute`, body),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["reservations"] });
      qc.invalidateQueries({ queryKey: ["materials"] });
      qc.invalidateQueries({ queryKey: ["lots"] });
    },
  });
}

export function useCancelReservation() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (id: number) => api.post<Reservation>(`/reservations/${id}/cancel`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["reservations"] }),
  });
}

export function useDeleteReservation() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (id: number) => api.delete(`/reservations/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["reservations"] }),
  });
}
