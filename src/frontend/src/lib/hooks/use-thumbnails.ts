import { useQuery } from "@tanstack/react-query";
import { getThumbnails } from "@/lib/api/thumbnails";

export function useThumbnails(driveId?: string, itemId?: string) {
  return useQuery({
    queryKey: ["thumbnails", driveId, itemId],
    queryFn: () => getThumbnails(driveId!, itemId!),
    enabled: Boolean(driveId && itemId),
    staleTime: 60_000,
  });
}
