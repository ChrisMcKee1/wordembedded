import { apiFetch } from "@/lib/api-client";
import type { DriveChildrenResponse, DriveItem } from "./types";

const enc = encodeURIComponent;

export function getRootChildren(driveId: string) {
  return apiFetch<DriveChildrenResponse>(`/api/drives/${enc(driveId)}/root/children`);
}

export function getChildren(driveId: string, itemId: string) {
  return apiFetch<DriveChildrenResponse>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}/children`);
}

export function getItem(driveId: string, itemId: string) {
  return apiFetch<DriveItem>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}`);
}

export function renameItem(driveId: string, itemId: string, name: string) {
  return apiFetch<DriveItem>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name }),
  });
}

export function deleteItem(driveId: string, itemId: string) {
  return apiFetch<void>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}`, { method: "DELETE" });
}

export function createFolder(driveId: string, parentId: string, name: string) {
  return apiFetch<DriveItem>(`/api/drives/${enc(driveId)}/items/${enc(parentId)}/folders`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name }),
  });
}
