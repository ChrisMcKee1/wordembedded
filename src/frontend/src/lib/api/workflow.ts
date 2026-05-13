import { apiFetch } from "@/lib/api-client";

export type WorkflowStatus = "Draft" | "InReview" | "Approved" | "Published";

export interface PublishedReference {
  driveId?: string | null;
  itemId?: string | null;
  webUrl?: string | null;
  name?: string | null;
  publishedAt?: string | null;
}

export interface ImportRequest {
  sourceDriveId: string;
  sourceItemId: string;
  targetDriveId: string;
  targetFolderId?: string;
  newName?: string;
}

export interface PublishRequest {
  targetDriveId: string;
  targetFolderId?: string;
  newName?: string;
}

export interface PublishTarget {
  siteId?: string | null;
  driveId?: string | null;
  folderId?: string | null;
  folderPath?: string | null;
}

export function importFromSharePoint(body: ImportRequest) {
  return apiFetch<{ monitorUrl?: string; status?: string }>(`/api/sharepoint/import`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
}

export function publishToSharePoint(driveId: string, itemId: string, body: PublishRequest) {
  return apiFetch<{ monitorUrl?: string; status?: string; published?: PublishedReference }>(`/api/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/publish`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
}

export function getWorkflowStatus(driveId: string, itemId: string) {
  return apiFetch<{ status: WorkflowStatus; published?: PublishedReference }>(`/api/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workflow`);
}

export function setWorkflowStatus(driveId: string, itemId: string, status: WorkflowStatus) {
  return apiFetch<{ status: WorkflowStatus }>(`/api/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workflow`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ status }),
  });
}

export function getPublishTarget(containerId: string) {
  return apiFetch<PublishTarget>(`/api/containers/${encodeURIComponent(containerId)}/publish-target`);
}

export function setPublishTarget(containerId: string, target: PublishTarget) {
  return apiFetch<PublishTarget>(`/api/containers/${encodeURIComponent(containerId)}/publish-target`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(target),
  });
}
