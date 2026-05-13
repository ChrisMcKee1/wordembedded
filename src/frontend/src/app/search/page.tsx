"use client";

import { Button, Card, Spinner, Text, makeStyles, tokens } from "@fluentui/react-components";
import { useRouter, useSearchParams } from "next/navigation";
import { AuthBoundary } from "@/components/auth/AuthBoundary";
import { AppShell } from "@/components/layout/AppShell";
import { FileIcon } from "@/components/files/FileIcon";
import { useSearch } from "@/lib/hooks/use-search";
import { useContainers } from "@/lib/hooks/use-containers";
import { formatDate } from "@/lib/files";

const useStyles = makeStyles({
  stack: { display: "grid", gap: tokens.spacingVerticalL },
  results: { display: "grid", gap: tokens.spacingVerticalS },
  result: { padding: tokens.spacingHorizontalM, display: "grid", gridTemplateColumns: "auto 1fr auto", gap: tokens.spacingHorizontalM, alignItems: "center", "@media (max-width: 720px)": { gridTemplateColumns: "auto 1fr" } },
  meta: { display: "grid", gap: tokens.spacingVerticalXXS },
  snippet: { color: tokens.colorNeutralForeground2 },
});

export default function SearchPage() {
  const styles = useStyles();
  const router = useRouter();
  const params = useSearchParams();
  const q = params.get("q") ?? "";
  const driveId = params.get("driveId") ?? undefined;
  const search = useSearch(q, driveId);
  const containers = useContainers();

  const containerForDrive = (hitDriveId: string) => containers.data?.find((container) => container.driveId === hitDriveId);

  return (
    <AuthBoundary>
      <AppShell>
        <section className={styles.stack}>
          <div>
            <Text as="h1" size={800} weight="semibold">Search</Text>
            <Text block>{q ? `Results for “${q}”` : "Enter a query in the search box."}</Text>
          </div>
          {search.isLoading ? <Spinner label="Searching" /> : null}
          <div className={styles.results}>
            {search.data?.value.map((hit) => {
              const container = hit.containerId ? undefined : containerForDrive(hit.driveId);
              const containerId = hit.containerId ?? container?.id;
              return (
                <Card key={`${hit.driveId}-${hit.id}`} className={styles.result} appearance="outline">
                  <FileIcon name={hit.name} mimeType={hit.file?.mimeType} item={hit.folder ? { name: hit.name, folder: { childCount: hit.folder.childCount ?? 0 } } : undefined} />
                  <div className={styles.meta}>
                    <Text weight="semibold">{hit.name}</Text>
                    <Text size={200}>{hit.containerName ?? container?.displayName ?? "Container"} {hit.parentReference?.path ? `· ${hit.parentReference.path}` : ""}</Text>
                    {hit.snippet ? <Text size={200} className={styles.snippet}>{hit.snippet}</Text> : null}
                    {hit.lastModifiedDateTime ? <Text size={200}>Modified {formatDate(hit.lastModifiedDateTime)}</Text> : null}
                  </div>
                  <Button disabled={!containerId || Boolean(hit.folder)} onClick={() => containerId && router.push(`/files/${encodeURIComponent(containerId)}/preview/${encodeURIComponent(hit.id)}?driveId=${encodeURIComponent(hit.driveId)}`)}>Open</Button>
                </Card>
              );
            })}
            {q && search.data?.value.length === 0 ? <Text>No results found.</Text> : null}
          </div>
        </section>
      </AppShell>
    </AuthBoundary>
  );
}
