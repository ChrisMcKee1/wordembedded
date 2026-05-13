"use client";

import { Button, Card, Input, Spinner, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowImportRegular, DocumentSearchRegular, SearchRegular } from "@fluentui/react-icons";
import { useEffect, useState } from "react";
import { useImportFromSharePoint, useSharePointSearch } from "@/lib/hooks/use-workflow";
import { useAppToast } from "@/components/common/Toaster";
import { FileIcon } from "@/components/files/FileIcon";
import { formatBytes, formatDate } from "@/lib/files";

const useStyles = makeStyles({
  wrapper: { display: "grid", gap: tokens.spacingVerticalS, padding: tokens.spacingHorizontalM },
  searchRow: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center" },
  list: { display: "grid", gap: tokens.spacingVerticalXS, maxHeight: "320px", overflow: "auto" },
  row: { display: "grid", gridTemplateColumns: "auto 1fr auto", gap: tokens.spacingHorizontalM, alignItems: "center", padding: tokens.spacingVerticalS, borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
  meta: { display: "grid", gap: tokens.spacingVerticalXXS, color: tokens.colorNeutralForeground2, minWidth: 0 },
  truncate: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  empty: { padding: tokens.spacingVerticalM, color: tokens.colorNeutralForeground2 },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center" },
});

export function SharePointSearchPanel({
  targetDriveId,
  targetFolderId,
  defaultExpanded = false,
}: {
  targetDriveId?: string;
  targetFolderId?: string;
  defaultExpanded?: boolean;
}) {
  const styles = useStyles();
  const notify = useAppToast();
  const [query, setQuery] = useState("");
  const [debounced, setDebounced] = useState("");
  const [expanded, setExpanded] = useState(defaultExpanded);
  const search = useSharePointSearch(debounced);
  const importMutation = useImportFromSharePoint(targetDriveId);

  useEffect(() => {
    const handle = setTimeout(() => setDebounced(query.trim()), 350);
    return () => clearTimeout(handle);
  }, [query]);

  const onImport = async (hit: { driveId?: string; id?: string; name?: string }) => {
    if (!targetDriveId || !hit.driveId || !hit.id) {
      notify({ title: "Cannot import", body: "Missing source drive, source item, or target drive.", intent: "error" });
      return;
    }
    try {
      const result = await importMutation.mutateAsync({
        sourceDriveId: hit.driveId,
        sourceItemId: hit.id,
        targetDriveId,
        targetFolderId,
      });
      notify({
        title: `Importing “${hit.name ?? "file"}”`,
        body: result.monitorUrl ? "Graph accepted the copy. It will appear shortly." : "Import queued.",
        intent: "success",
      });
    } catch (error) {
      notify({ title: "Import failed", body: error instanceof Error ? error.message : "Unknown error", intent: "error" });
    }
  };

  return (
    <Card className={styles.wrapper} appearance="outline">
      <div className={styles.header}>
        <Text weight="semibold">Pull files from SharePoint or OneDrive</Text>
        <Button appearance="subtle" onClick={() => setExpanded((v) => !v)}>
          {expanded ? "Hide" : "Show"}
        </Button>
      </div>
      {expanded ? (
        <>
          <div className={styles.searchRow}>
            <Input
              placeholder="Find a file you have access to in M365…"
              value={query}
              onChange={(_, data) => setQuery(data.value)}
              contentBefore={<SearchRegular />}
              style={{ flex: 1 }}
            />
          </div>
          {search.isFetching ? <Spinner size="tiny" label="Searching" /> : null}
          {!search.isFetching && debounced.length > 1 && (search.data?.value.length ?? 0) === 0 ? (
            <div className={styles.empty}><DocumentSearchRegular /> <Text>No matches.</Text></div>
          ) : null}
          {search.data && search.data.value.length > 0 ? (
            <div className={styles.list} aria-label="SharePoint search results">
              {search.data.value.map((hit) => (
                <div key={hit.hitId ?? hit.id} className={styles.row}>
                  <FileIcon name={hit.name ?? ""} />
                  <div className={styles.meta}>
                    <Text className={styles.truncate} weight="semibold" title={hit.name ?? ""}>{hit.name ?? "Untitled"}</Text>
                    <Text className={styles.truncate} size={200} title={hit.parentPath ?? ""}>{hit.parentPath ?? ""}</Text>
                    <Text size={200}>{formatBytes(hit.size ?? 0)} • {hit.lastModifiedDateTime ? formatDate(hit.lastModifiedDateTime) : ""}</Text>
                  </div>
                  <Button
                    appearance="primary"
                    icon={<ArrowImportRegular />}
                    size="small"
                    disabled={!targetDriveId || importMutation.isPending}
                    onClick={() => void onImport(hit)}
                  >
                    Import
                  </Button>
                </div>
              ))}
            </div>
          ) : null}
        </>
      ) : null}
    </Card>
  );
}
