import { useMutation, useQueryClient } from "@tanstack/react-query";
import type { DriveChildrenResponse, DriveItem } from "@/lib/api/types";
import { createFolder } from "@/lib/api/files";
import { childrenKey } from "./use-children";

interface CreateFolderVariables {
  driveId: string;
  parentId: string;
  name: string;
}

export function useCreateFolder() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ driveId, parentId, name }: CreateFolderVariables) => createFolder(driveId, parentId, name),
    onMutate: async (vars) => {
      await queryClient.cancelQueries({ queryKey: childrenKey(vars.driveId, vars.parentId) });
      const previousChildren = queryClient.getQueryData<DriveChildrenResponse>(childrenKey(vars.driveId, vars.parentId));
      const optimistic: DriveItem = {
        id: `pending-${Date.now()}`,
        name: vars.name,
        webUrl: "",
        size: 0,
        createdDateTime: new Date().toISOString(),
        lastModifiedDateTime: new Date().toISOString(),
        folder: { childCount: 0 },
        parentReference: { id: vars.parentId, path: "", driveId: vars.driveId },
      };
      queryClient.setQueryData<DriveChildrenResponse>(childrenKey(vars.driveId, vars.parentId), (current) =>
        current ? { value: [optimistic, ...current.value] } : { value: [optimistic] },
      );
      return { previousChildren };
    },
    onError: (_error, vars, context) => {
      if (context?.previousChildren) queryClient.setQueryData(childrenKey(vars.driveId, vars.parentId), context.previousChildren);
    },
    onSettled: (_data, _error, vars) => {
      void queryClient.invalidateQueries({ queryKey: childrenKey(vars.driveId, vars.parentId) });
    },
  });
}
