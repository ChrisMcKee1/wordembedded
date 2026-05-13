import { useMutation, useQueryClient } from "@tanstack/react-query";
import { revokePermission } from "@/lib/api/sharing";

interface RevokeVariables {
  driveId: string;
  itemId: string;
  permissionId: string;
}

export function useRevokePermission() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ driveId, itemId, permissionId }: RevokeVariables) => revokePermission(driveId, itemId, permissionId),
    onSuccess: (_data, vars) => void queryClient.invalidateQueries({ queryKey: ["permissions", vars.driveId, vars.itemId] }),
  });
}
