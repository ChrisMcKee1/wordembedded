import { useMutation, useQueryClient } from "@tanstack/react-query";
import { createLink } from "@/lib/api/sharing";
import type { CreateLinkRequest } from "@/lib/api/types";

interface CreateLinkVariables {
  driveId: string;
  itemId: string;
  request: CreateLinkRequest;
}

export function useCreateLink() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ driveId, itemId, request }: CreateLinkVariables) => createLink(driveId, itemId, request),
    onSuccess: (_data, vars) => void queryClient.invalidateQueries({ queryKey: ["permissions", vars.driveId, vars.itemId] }),
  });
}
