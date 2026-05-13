import { useMutation, useQueryClient } from "@tanstack/react-query";
import { invitePeople } from "@/lib/api/sharing";
import type { InviteRequest } from "@/lib/api/types";

interface InviteVariables {
  driveId: string;
  itemId: string;
  request: InviteRequest;
}

export function useInvite() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ driveId, itemId, request }: InviteVariables) => invitePeople(driveId, itemId, request),
    onSuccess: (_data, vars) => void queryClient.invalidateQueries({ queryKey: ["permissions", vars.driveId, vars.itemId] }),
  });
}
