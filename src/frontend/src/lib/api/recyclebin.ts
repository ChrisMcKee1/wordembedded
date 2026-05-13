import { apiFetch } from "@/lib/api-client";
import type { RecycleBinResponse } from "./types";

const enc = encodeURIComponent;

export function listRecycleBin(containerId: string) {
  return apiFetch<RecycleBinResponse>(`/api/containers/${enc(containerId)}/recyclebin`);
}

export function restoreRecycleBin(containerId: string, ids: string[]) {
  return apiFetch<{ restored: string[] }>(`/api/containers/${enc(containerId)}/recyclebin/restore`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids }),
  });
}

export function purgeRecycleBin(containerId: string, ids: string[]) {
  return apiFetch<void>(`/api/containers/${enc(containerId)}/recyclebin/items`, {
    method: "DELETE",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ids }),
  });
}
