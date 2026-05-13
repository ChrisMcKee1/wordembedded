import { useMutation, useQueryClient } from "@tanstack/react-query";
import { restoreVersion } from "@/lib/api/versions";

interface RestoreVariables {
  driveId: string;
  itemId: string;
  versionId: string;
}

export function useRestoreVersion() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ driveId, itemId, versionId }: RestoreVariables) => restoreVersion(driveId, itemId, versionId),
    onSuccess: (_data, vars) => {
      void queryClient.invalidateQueries({ queryKey: ["versions", vars.driveId, vars.itemId] });
      void queryClient.invalidateQueries({ queryKey: ["item", vars.driveId, vars.itemId] });
    },
  });
}
