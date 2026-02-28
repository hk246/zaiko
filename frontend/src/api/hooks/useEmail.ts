import { useMutation } from "@tanstack/react-query";
import { api } from "../client";

interface EmailDraftRequest {
  material_id: number;
  contact_id: number;
}

interface EmailDraftResponse {
  status: string;
  message: string;
}

export function useEmailDraft() {
  return useMutation({
    mutationFn: (body: EmailDraftRequest) =>
      api.post<EmailDraftResponse>("/email/draft", body),
  });
}
