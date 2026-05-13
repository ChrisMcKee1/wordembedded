import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { listRecycleBin, purgeRecycleBin, restoreRecycleBin } from "@/lib/api/recyclebin";
import { childrenKey } from "./use-children";

export function useRecycleBin(containerId?: string, enabled = true) {
  return useQuery({
    queryKey: ["recyclebin", containerId],
    queryFn: () => listRecycleBin(containerId!),
    enabled: Boolean(containerId) && enabled,
    staleTime: 0,
  });
}

export function useRestoreRecycleBin(containerId: string | undefined, driveId: string | undefined) {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (ids: string[]) => restoreRecycleBin(containerId!, ids),
    onSuccess: () => {
      void queryClient.invalidateQueries({ queryKey: ["recyclebin", containerId] });
      if (driveId) void queryClient.invalidateQueries({ queryKey: childrenKey(driveId) });
    },
  });
}

export function usePurgeRecycleBin(containerId: string | undefined) {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (ids: string[]) => purgeRecycleBin(containerId!, ids),
    onSuccess: () => void queryClient.invalidateQueries({ queryKey: ["recyclebin", containerId] }),
  });
}
