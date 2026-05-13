import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import {
  getPublishTarget,
  getWorkflowStatus,
  importFromSharePoint,
  publishToSharePoint,
  setPublishTarget,
  setWorkflowStatus,
  type ImportRequest,
  type PublishRequest,
  type PublishTarget,
  type WorkflowStatus,
} from "@/lib/api/workflow";
import { listSharePointChildren, listSharePointDrives, listSharePointSites, searchSharePoint } from "@/lib/api/sharepoint";
import { childrenKey } from "./use-children";

export function useSharePointSearch(query: string) {
  return useQuery({
    queryKey: ["sharepoint", "search", query],
    queryFn: () => searchSharePoint(query),
    enabled: query.trim().length > 1,
    staleTime: 30_000,
  });
}

export function useSharePointSites(query?: string) {
  return useQuery({
    queryKey: ["sharepoint", "sites", query ?? ""],
    queryFn: () => listSharePointSites(query),
    staleTime: 60_000,
  });
}

export function useSharePointDrives(siteId?: string) {
  return useQuery({
    queryKey: ["sharepoint", "drives", siteId],
    queryFn: () => listSharePointDrives(siteId!),
    enabled: Boolean(siteId),
    staleTime: 60_000,
  });
}

export function useSharePointChildren(driveId?: string, itemId?: string) {
  return useQuery({
    queryKey: ["sharepoint", "children", driveId, itemId ?? "root"],
    queryFn: () => listSharePointChildren(driveId!, itemId),
    enabled: Boolean(driveId),
    staleTime: 30_000,
  });
}

export function useImportFromSharePoint(targetDriveId?: string) {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (body: ImportRequest) => importFromSharePoint(body),
    onSuccess: () => {
      if (targetDriveId) void queryClient.invalidateQueries({ queryKey: childrenKey(targetDriveId) });
    },
  });
}

export function usePublishToSharePoint(driveId?: string) {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: ({ itemId, ...body }: PublishRequest & { itemId: string }) => publishToSharePoint(driveId!, itemId, body),
    onSuccess: (_data, vars) => {
      if (driveId) {
        void queryClient.invalidateQueries({ queryKey: ["workflow", driveId, vars.itemId] });
        void queryClient.invalidateQueries({ queryKey: childrenKey(driveId) });
      }
    },
  });
}

export function useWorkflowStatus(driveId?: string, itemId?: string) {
  return useQuery({
    queryKey: ["workflow", driveId, itemId],
    queryFn: () => getWorkflowStatus(driveId!, itemId!),
    enabled: Boolean(driveId && itemId),
    staleTime: 10_000,
  });
}

export function useSetWorkflowStatus(driveId?: string, itemId?: string) {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (status: WorkflowStatus) => setWorkflowStatus(driveId!, itemId!, status),
    onSuccess: () => {
      if (driveId && itemId) void queryClient.invalidateQueries({ queryKey: ["workflow", driveId, itemId] });
    },
  });
}

export function usePublishTarget(containerId?: string) {
  return useQuery({
    queryKey: ["publish-target", containerId],
    queryFn: () => getPublishTarget(containerId!),
    enabled: Boolean(containerId),
    staleTime: 60_000,
  });
}

export function useSetPublishTarget(containerId?: string) {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: (target: PublishTarget) => setPublishTarget(containerId!, target),
    onSuccess: () => {
      if (containerId) void queryClient.invalidateQueries({ queryKey: ["publish-target", containerId] });
    },
  });
}
