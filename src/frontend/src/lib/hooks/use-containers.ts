import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { createContainer, getContainer, listContainers, type CreateContainerInput } from "@/lib/api/containers";
import type { Container } from "@/lib/api/types";

export function useContainers() {
  return useQuery({ queryKey: ["containers"], queryFn: listContainers });
}

export function useContainer(containerId?: string) {
  return useQuery({ queryKey: ["container", containerId], queryFn: () => getContainer(containerId!), enabled: Boolean(containerId) });
}

export function useCreateContainer() {
  const queryClient = useQueryClient();
  return useMutation<Container, Error, CreateContainerInput>({
    mutationFn: createContainer,
    onSuccess: (created) => {
      queryClient.setQueryData<Container[] | undefined>(["containers"], (existing) =>
        existing ? [created, ...existing.filter((c) => c.id !== created.id)] : [created],
      );
      queryClient.invalidateQueries({ queryKey: ["containers"] });
    },
  });
}

