import type { DriveItem, ModifiedBy } from "@/lib/api/types";

export interface RoutePathSegment {
  id: string;
  name: string;
}

export function formatBytes(size?: number) {
  if (!size) return "—";
  const units = ["B", "KB", "MB", "GB", "TB"];
  let value = size;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit += 1;
  }
  return `${value >= 10 || unit === 0 ? value.toFixed(0) : value.toFixed(1)} ${units[unit]}`;
}

export function formatDate(value?: string) {
  if (!value) return "—";
  return new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short" }).format(new Date(value));
}

export function modifiedByName(modifiedBy?: ModifiedBy) {
  return modifiedBy?.user?.displayName ?? modifiedBy?.displayName ?? modifiedBy?.user?.email ?? modifiedBy?.email ?? "Unknown";
}

export function modifiedByEmail(modifiedBy?: ModifiedBy) {
  return modifiedBy?.user?.email ?? modifiedBy?.email;
}

export function extensionFor(name: string) {
  const last = name.split(".").pop();
  return last && last !== name ? last.toLowerCase() : "";
}

export function officeAppName(item: Pick<DriveItem, "name" | "file">) {
  const ext = extensionFor(item.name);
  const mime = item.file?.mimeType.toLowerCase() ?? "";
  if (["doc", "docx", "dotx"].includes(ext) || mime.includes("word")) return "Word";
  if (["xls", "xlsx", "xlsm", "csv"].includes(ext) || mime.includes("excel") || mime.includes("spreadsheet")) return "Excel";
  if (["ppt", "pptx"].includes(ext) || mime.includes("powerpoint") || mime.includes("presentation")) return "PowerPoint";
  return "Office";
}

export function supportsOfficeEdit(item: Pick<DriveItem, "name" | "file" | "folder">) {
  if (item.folder) return false;
  return ["Word", "Excel", "PowerPoint"].includes(officeAppName(item));
}

export function encodeRouteSegment(segment: RoutePathSegment) {
  return `${encodeURIComponent(segment.id)}--${encodeURIComponent(segment.name)}`;
}

export function decodeRouteSegment(value: string): RoutePathSegment {
  const [id, ...nameParts] = value.split("--");
  return { id: decodeURIComponent(id), name: decodeURIComponent(nameParts.join("--") || "Folder") };
}

export function previewUrlWithNoBanner(url?: string) {
  if (!url) return undefined;
  return url.includes("?") ? `${url}&nb=true` : `${url}?nb=true`;
}
