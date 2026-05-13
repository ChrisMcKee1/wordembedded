import { apiFetch } from "@/lib/api-client";
import type { SearchResponse } from "./types";

export function searchFiles(query: string, driveId?: string) {
  const params = new URLSearchParams({ q: query });
  if (driveId) params.set("driveId", driveId);
  return apiFetch<SearchResponse>(`/api/search?${params.toString()}`);
}
