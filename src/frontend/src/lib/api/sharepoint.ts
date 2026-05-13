import { apiFetch } from "@/lib/api-client";

export interface SharePointSearchHit {
  hitId?: string;
  id?: string;
  name?: string;
  webUrl?: string;
  size?: number;
  lastModifiedDateTime?: string;
  driveId?: string;
  parentPath?: string;
  siteId?: string;
  summary?: string;
}

export interface SharePointSiteSummary {
  id?: string;
  displayName?: string;
  name?: string;
  webUrl?: string;
}

export interface SharePointDriveSummary {
  id?: string;
  name?: string;
  webUrl?: string;
  driveType?: string;
}

export interface SharePointBrowseItem {
  id?: string;
  name?: string;
  webUrl?: string;
  size?: number;
  lastModifiedDateTime?: string;
  isFolder: boolean;
  mimeType?: string;
}

export function searchSharePoint(q: string) {
  return apiFetch<{ value: SharePointSearchHit[] }>(`/api/sharepoint/search?q=${encodeURIComponent(q)}`);
}

export function listSharePointSites(q?: string) {
  const qs = q ? `?q=${encodeURIComponent(q)}` : "";
  return apiFetch<{ value: SharePointSiteSummary[] }>(`/api/sharepoint/sites${qs}`);
}

export function listSharePointDrives(siteId: string) {
  return apiFetch<{ value: SharePointDriveSummary[] }>(`/api/sharepoint/sites/${encodeURIComponent(siteId)}/drives`);
}

export function listSharePointChildren(driveId: string, itemId?: string) {
  const qs = itemId ? `?itemId=${encodeURIComponent(itemId)}` : "";
  return apiFetch<{ value: SharePointBrowseItem[] }>(`/api/sharepoint/drives/${encodeURIComponent(driveId)}/children${qs}`);
}
