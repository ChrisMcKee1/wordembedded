import { useMutation, useQueryClient } from "@tanstack/react-query";
import type { DriveChildrenResponse } from "@/lib/api/types";
import { deleteItem } from "@/lib/api/files";
import { childrenKey } from "./use-children";

interface DeleteVariables {
  driveId: string;
  itemId: string;
  parentId?: string;
}

export function useDelete() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ driveId, itemId }: DeleteVariables) => deleteItem(driveId, itemId),
    onMutate: async (vars) => {
      await queryClient.cancelQueries({ queryKey: childrenKey(vars.driveId, vars.parentId) });
      const previousChildren = queryClient.getQueryData<DriveChildrenResponse>(childrenKey(vars.driveId, vars.parentId));
      queryClient.setQueryData<DriveChildrenResponse>(childrenKey(vars.driveId, vars.parentId), (current) =>
        current ? { value: current.value.filter((item) => item.id !== vars.itemId) } : current,
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
