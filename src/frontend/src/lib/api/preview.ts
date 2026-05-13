import { apiFetch } from "@/lib/api-client";
import type { PreviewResponse } from "./types";

const enc = encodeURIComponent;

export interface PreviewOptions {
  allowEdit?: boolean;
  viewer?: "office" | "onedrive";
  chromeless?: boolean;
  page?: string;
  zoom?: number;
}

export function getPreview(driveId: string, itemId: string, options?: PreviewOptions) {
  const hasBody = options && Object.keys(options).length > 0;
  return apiFetch<PreviewResponse>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}/preview`, {
    method: "POST",
    headers: hasBody ? { "Content-Type": "application/json" } : undefined,
    body: hasBody ? JSON.stringify(options) : undefined,
  });
}

