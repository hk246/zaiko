import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { api } from "../client";
import type { Material, MaterialListItem, Lot } from "../types";

export function useMaterials(params?: { q?: string; label_id?: number; low_stock_only?: boolean }) {
  const search = new URLSearchParams();
  if (params?.q) search.set("q", params.q);
  if (params?.label_id) search.set("label_id", String(params.label_id));
  if (params?.low_stock_only) search.set("low_stock_only", "true");
  const qs = search.toString();
  return useQuery({
    queryKey: ["materials", qs],
    queryFn: () => api.get<MaterialListItem[]>(`/materials${qs ? `?${qs}` : ""}`),
  });
}

export function useMaterial(id: number | string) {
  return useQuery({
    queryKey: ["material", id],
    queryFn: () => api.get<Material>(`/materials/${id}`),
    enabled: !!id,
  });
}

export function useLots(materialId: number | string) {
  return useQuery({
    queryKey: ["lots", materialId],
    queryFn: () => api.get<Lot[]>(`/materials/${materialId}/lots`),
    enabled: !!materialId,
  });
}

export function useCreateMaterial() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (body: Record<string, unknown>) => api.post<Material>("/materials", body),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["materials"] }),
  });
}

export function useUpdateMaterial(id: number | string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (body: Record<string, unknown>) => api.put<Material>(`/materials/${id}`, body),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["materials"] });
      qc.invalidateQueries({ queryKey: ["material", id] });
    },
  });
}

export function useDeleteMaterial() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (id: number) => api.delete(`/materials/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["materials"] }),
  });
}

export function useCreateLot(materialId: number | string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (body: Record<string, unknown>) =>
      api.post<Lot>(`/materials/${materialId}/lots`, body),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["lots", materialId] });
      qc.invalidateQueries({ queryKey: ["material", materialId] });
      qc.invalidateQueries({ queryKey: ["materials"] });
    },
  });
}

export function useUpdateLot(materialId: number | string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: ({ lotId, body }: { lotId: number; body: Record<string, unknown> }) =>
      api.put<Lot>(`/materials/${materialId}/lots/${lotId}`, body),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["lots", materialId] });
      qc.invalidateQueries({ queryKey: ["material", materialId] });
      qc.invalidateQueries({ queryKey: ["materials"] });
    },
  });
}

export function useDeleteLot(materialId: number | string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (lotId: number) => api.delete(`/materials/${materialId}/lots/${lotId}`),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["lots", materialId] });
      qc.invalidateQueries({ queryKey: ["material", materialId] });
      qc.invalidateQueries({ queryKey: ["materials"] });
    },
  });
}

export function useToggleFraction(materialId: number | string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (lotId: number) =>
      api.patch<Lot>(`/materials/${materialId}/lots/${lotId}/fraction`),
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ["lots", materialId] });
      qc.invalidateQueries({ queryKey: ["material", materialId] });
    },
  });
}
