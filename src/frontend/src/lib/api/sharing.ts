import { apiFetch } from "@/lib/api-client";
import type { CreateLinkRequest, CreateLinkResponse, InviteRequest, Permission } from "./types";

const enc = encodeURIComponent;

export function createLink(driveId: string, itemId: string, request: CreateLinkRequest) {
  return apiFetch<CreateLinkResponse>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}/createLink`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
  });
}

export function invitePeople(driveId: string, itemId: string, request: InviteRequest) {
  return apiFetch<Permission[]>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}/invite`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(request),
  });
}

export function listPermissions(driveId: string, itemId: string) {
  return apiFetch<{ value: Permission[] }>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}/permissions`);
}

export function revokePermission(driveId: string, itemId: string, permissionId: string) {
  return apiFetch<void>(`/api/drives/${enc(driveId)}/items/${enc(itemId)}/permissions/${enc(permissionId)}`, { method: "DELETE" });
}
