import { useQuery } from "@tanstack/react-query";
import { listVersions } from "@/lib/api/versions";

export function useVersions(driveId?: string, itemId?: string) {
  return useQuery({ queryKey: ["versions", driveId, itemId], queryFn: () => listVersions(driveId!, itemId!), enabled: Boolean(driveId && itemId) });
}
