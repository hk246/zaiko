import { useQuery } from "@tanstack/react-query";
import { api } from "../client";
import type { Contact } from "../types";

export function useContacts() {
  return useQuery({
    queryKey: ["contacts"],
    queryFn: () => api.get<Contact[]>("/contacts"),
  });
}
