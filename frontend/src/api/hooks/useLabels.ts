import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { api } from "../client";
import type { Label } from "../types";

export function useLabels() {
  return useQuery({
    queryKey: ["labels"],
    queryFn: () => api.get<Label[]>("/labels"),
  });
}

export function useCreateLabel() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (body: { name: string; color: string }) => api.post<Label>("/labels", body),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["labels"] }),
  });
}

export function useUpdateLabel(id: number) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (body: { name?: string; color?: string }) => api.put<Label>(`/labels/${id}`, body),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["labels"] }),
  });
}

export function useDeleteLabel() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: (id: number) => api.delete(`/labels/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ["labels"] }),
  });
}
