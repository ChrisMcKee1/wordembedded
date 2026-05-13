"use client";

import { useThumbnails } from "@/lib/hooks/use-thumbnails";
import { FileIcon } from "@/components/files/FileIcon";
import type { DriveItem } from "@/lib/api/types";

export function ThumbnailTile({ driveId, item, size = "medium" }: { driveId?: string; item: DriveItem; size?: "small" | "medium" | "large" }) {
  const thumbs = useThumbnails(driveId, item.id);
  const set = thumbs.data?.value?.[0];
  const url = set?.[size]?.url ?? set?.medium?.url ?? set?.large?.url ?? set?.small?.url;

  if (item.folder || !url) {
    return <div style={{ display: "grid", placeItems: "center", height: 96, fontSize: 32 }}><FileIcon item={item} /></div>;
  }

  return (
    <img
      src={url}
      alt={item.name}
      loading="lazy"
      decoding="async"
      style={{ width: "100%", height: 96, objectFit: "cover", borderRadius: 4, background: "var(--colorNeutralBackground2)" }}
    />
  );
}
