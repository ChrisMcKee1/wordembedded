import { useQuery } from "@tanstack/react-query";
import { getItem } from "@/lib/api/files";

export function useItem(driveId?: string, itemId?: string) {
  return useQuery({ queryKey: ["item", driveId, itemId], queryFn: () => getItem(driveId!, itemId!), enabled: Boolean(driveId && itemId) });
}
