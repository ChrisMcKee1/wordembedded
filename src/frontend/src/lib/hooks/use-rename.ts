import { useMutation, useQueryClient } from "@tanstack/react-query";
import type { DriveChildrenResponse, DriveItem } from "@/lib/api/types";
import { renameItem } from "@/lib/api/files";
import { childrenKey } from "./use-children";

interface RenameVariables {
  driveId: string;
  itemId: string;
  parentId?: string;
  name: string;
}

export function useRename() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ driveId, itemId, name }: RenameVariables) => renameItem(driveId, itemId, name),
    onMutate: async (vars) => {
      await queryClient.cancelQueries({ queryKey: childrenKey(vars.driveId, vars.parentId) });
      const previousChildren = queryClient.getQueryData<DriveChildrenResponse>(childrenKey(vars.driveId, vars.parentId));
      const previousItem = queryClient.getQueryData<DriveItem>(["item", vars.driveId, vars.itemId]);
      queryClient.setQueryData<DriveChildrenResponse>(childrenKey(vars.driveId, vars.parentId), (current) =>
        current ? { value: current.value.map((item) => (item.id === vars.itemId ? { ...item, name: vars.name } : item)) } : current,
      );
      queryClient.setQueryData<DriveItem>(["item", vars.driveId, vars.itemId], (current) => (current ? { ...current, name: vars.name } : current));
      return { previousChildren, previousItem };
    },
    onError: (_error, vars, context) => {
      if (context?.previousChildren) queryClient.setQueryData(childrenKey(vars.driveId, vars.parentId), context.previousChildren);
      if (context?.previousItem) queryClient.setQueryData(["item", vars.driveId, vars.itemId], context.previousItem);
    },
    onSettled: (_data, _error, vars) => {
      void queryClient.invalidateQueries({ queryKey: childrenKey(vars.driveId, vars.parentId) });
      void queryClient.invalidateQueries({ queryKey: ["item", vars.driveId, vars.itemId] });
    },
  });
}
