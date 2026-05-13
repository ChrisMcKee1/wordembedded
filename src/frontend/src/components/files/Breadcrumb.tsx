"use client";

import { Breadcrumb, BreadcrumbButton, BreadcrumbDivider, BreadcrumbItem } from "@fluentui/react-components";
import { useRouter } from "next/navigation";
import type { RoutePathSegment } from "@/lib/files";
import { encodeRouteSegment } from "@/lib/files";

interface FileBreadcrumbProps {
  containerId: string;
  containerName?: string;
  segments: RoutePathSegment[];
}

export function FileBreadcrumb({ containerId, containerName = "Files", segments }: FileBreadcrumbProps) {
  const router = useRouter();

  const navigateTo = (index: number) => {
    if (index < 0) {
      router.push(`/files/${encodeURIComponent(containerId)}`);
      return;
    }
    const path = segments.slice(0, index + 1).map(encodeRouteSegment).join("/");
    router.push(`/files/${encodeURIComponent(containerId)}/${path}`);
  };

  return (
    <Breadcrumb aria-label="Folder path">
      <BreadcrumbItem>
        <BreadcrumbButton onClick={() => navigateTo(-1)} current={segments.length === 0}>{containerName}</BreadcrumbButton>
      </BreadcrumbItem>
      {segments.map((segment, index) => (
        <span key={`${segment.id}-${index}`}>
          <BreadcrumbDivider />
          <BreadcrumbItem>
            <BreadcrumbButton onClick={() => navigateTo(index)} current={index === segments.length - 1}>{segment.name}</BreadcrumbButton>
          </BreadcrumbItem>
        </span>
      ))}
    </Breadcrumb>
  );
}
