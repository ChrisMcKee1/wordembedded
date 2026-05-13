import { useMutation, useQueryClient } from "@tanstack/react-query";
import { uploadFile } from "@/lib/api/upload";
import { childrenKey } from "./use-children";

interface UploadVariables {
  driveId: string;
  parentId: string;
  file: File;
  onProgress?: (loaded: number, total: number) => void;
}

export function useUpload() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ driveId, parentId, file, onProgress }: UploadVariables) => uploadFile(driveId, parentId, file, onProgress),
    onSettled: (_data, _error, vars) => {
      void queryClient.invalidateQueries({ queryKey: childrenKey(vars.driveId, vars.parentId) });
    },
  });
}
