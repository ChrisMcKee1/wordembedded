import { apiFetch } from "@/lib/api-client";
import type { DriveItem, UploadSessionResponse } from "./types";

const enc = encodeURIComponent;
const smallUploadLimit = 4 * 1024 * 1024;
const chunkSize = 320 * 1024 * 10;

export function isSmallUpload(file: File) {
  return file.size <= smallUploadLimit;
}

export function createUploadSession(driveId: string, parentId: string, file: File) {
  return apiFetch<UploadSessionResponse>(`/api/drives/${enc(driveId)}/items/${enc(parentId)}/upload-session`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: file.name, size: file.size }),
  });
}

export function uploadSmall(driveId: string, parentId: string, file: File) {
  const body = new FormData();
  body.append("file", file, file.name);
  return apiFetch<DriveItem>(`/api/drives/${enc(driveId)}/items/${enc(parentId)}/upload`, { method: "POST", body });
}

export async function uploadLarge(driveId: string, parentId: string, file: File, onProgress?: (loaded: number, total: number) => void) {
  const session = await createUploadSession(driveId, parentId, file);
  let start = 0;

  while (start < file.size) {
    const end = Math.min(start + chunkSize, file.size);
    const response = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: { "Content-Range": `bytes ${start}-${end - 1}/${file.size}` },
      body: file.slice(start, end),
    });

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(`${response.status} ${response.statusText}: ${text}`);
    }

    start = end;
    onProgress?.(start, file.size);

    if (response.status === 200 || response.status === 201) {
      return (await response.json()) as DriveItem;
    }
  }

  throw new Error("Upload session completed without a drive item response.");
}

export function uploadFile(driveId: string, parentId: string, file: File, onProgress?: (loaded: number, total: number) => void) {
  if (isSmallUpload(file)) {
    onProgress?.(0, file.size);
    return uploadSmall(driveId, parentId, file).then((item) => {
      onProgress?.(file.size, file.size);
      return item;
    });
  }

  return uploadLarge(driveId, parentId, file, onProgress);
}
