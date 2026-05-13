import { apiFetch } from "@/lib/api-client";
import type { ThumbnailSetResponse } from "./types";

const enc = encodeURIComponent;

export function getThumbnails(driveId: string, itemId: string) {
  return apiFetch<ThumbnailSetResponse>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}/thumbnails`);
}
