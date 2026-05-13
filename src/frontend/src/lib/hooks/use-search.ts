import { useQuery } from "@tanstack/react-query";
import { searchFiles } from "@/lib/api/search";

export function useSearch(query: string, driveId?: string) {
  return useQuery({ queryKey: ["search", query, driveId ?? "all"], queryFn: () => searchFiles(query, driveId), enabled: query.trim().length > 0 });
}
