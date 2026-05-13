import { useQuery } from "@tanstack/react-query";
import { getChildren, getRootChildren } from "@/lib/api/files";

export function childrenKey(driveId?: string, itemId?: string) {
  return ["children", driveId, itemId ?? "root"] as const;
}

export function useChildren(driveId?: string, itemId?: string) {
  return useQuery({
    queryKey: childrenKey(driveId, itemId),
    queryFn: () => (itemId ? getChildren(driveId!, itemId) : getRootChildren(driveId!)),
    enabled: Boolean(driveId),
  });
}
