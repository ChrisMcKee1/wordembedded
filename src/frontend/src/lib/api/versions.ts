import { apiFetch } from "@/lib/api-client";
import type { VersionsResponse } from "./types";

const enc = encodeURIComponent;

export function listVersions(driveId: string, itemId: string) {
  return apiFetch<VersionsResponse>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}/versions`);
}

export function restoreVersion(driveId: string, itemId: string, versionId: string) {
  return apiFetch<void>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}/versions/${enc(versionId)}/restore`, { method: "POST" });
}
