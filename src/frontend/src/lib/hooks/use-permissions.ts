import { useQuery } from "@tanstack/react-query";
import { listPermissions } from "@/lib/api/sharing";

export function usePermissions(driveId?: string, itemId?: string) {
  return useQuery({ queryKey: ["permissions", driveId, itemId], queryFn: () => listPermissions(driveId!, itemId!), enabled: Boolean(driveId && itemId) });
}
