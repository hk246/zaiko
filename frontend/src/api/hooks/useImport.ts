import { useMutation } from "@tanstack/react-query";
import { api } from "@/api/client";

export interface ImportPreview {
  total_rows: number;
  errors: { row: number; field: string; message: string }[];
  valid: boolean;
}

export interface ImportResult {
  imported: number;
  message: string;
}

export function useImportPreview() {
  return useMutation({
    mutationFn: ({ dataType, file, materialId }: { dataType: string; file: File; materialId?: number }) => {
      const params = new URLSearchParams({ confirm: "false" });
      if (materialId) params.set("material_id", String(materialId));
      return api.upload<ImportPreview>(`/import/${dataType}?${params}`, file);
    },
  });
}

export function useImportConfirm() {
  return useMutation({
    mutationFn: ({ dataType, file, materialId }: { dataType: string; file: File; materialId?: number }) => {
      const params = new URLSearchParams({ confirm: "true" });
      if (materialId) params.set("material_id", String(materialId));
      return api.upload<ImportResult>(`/import/${dataType}?${params}`, file);
    },
  });
}
