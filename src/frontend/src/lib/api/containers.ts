import { apiFetch } from "@/lib/api-client";
import type { Container } from "./types";

export function listContainers() {
  return apiFetch<Container[]>("/api/containers");
}

export function getContainer(containerId: string) {
  return apiFetch<Container>(`/api/containers/${encodeURIComponent(containerId)}`);
}

export interface CreateContainerInput {
  displayName: string;
  description?: string;
}

export function createContainer(input: CreateContainerInput) {
  return apiFetch<Container>("/api/containers", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });
}

