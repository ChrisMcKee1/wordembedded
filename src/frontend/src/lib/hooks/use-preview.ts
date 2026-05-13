import { useQuery } from "@tanstack/react-query";
import { getPreview, type PreviewOptions } from "@/lib/api/preview";

export function usePreview(driveId?: string, itemId?: string, options?: PreviewOptions) {
  const key = JSON.stringify(options ?? {});
  return useQuery({
    queryKey: ["preview", driveId, itemId, key],
    queryFn: () => getPreview(driveId!, itemId!, options),
    enabled: Boolean(driveId && itemId),
    staleTime: 0,
    retry: 1,
  });
}

