"use client";

import { Button, Checkbox, Drawer, DrawerBody, DrawerHeader, DrawerHeaderTitle, Spinner, Text, Toolbar, ToolbarButton, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowClockwiseRegular, DeleteRegular, DismissRegular } from "@fluentui/react-icons";
import { useMemo, useState } from "react";
import { useRecycleBin, useRestoreRecycleBin, usePurgeRecycleBin } from "@/lib/hooks/use-recyclebin";
import { useAppToast } from "@/components/common/Toaster";
import { formatBytes, formatDate } from "@/lib/files";

const useStyles = makeStyles({
  body: { display: "grid", gap: tokens.spacingVerticalM },
  toolbar: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center" },
  row: { display: "grid", gridTemplateColumns: "auto 1fr auto", gap: tokens.spacingHorizontalM, padding: tokens.spacingVerticalS, alignItems: "center", borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
  meta: { display: "grid", gap: tokens.spacingVerticalXXS, color: tokens.colorNeutralForeground2 },
  empty: { display: "grid", placeItems: "center", padding: tokens.spacingVerticalXXXL, color: tokens.colorNeutralForeground2 },
});

export function RecycleBinDrawer({
  containerId,
  driveId,
  open,
  onOpenChange,
}: {
  containerId: string;
  driveId?: string;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}) {
  const styles = useStyles();
  const notify = useAppToast();
  const bin = useRecycleBin(containerId, open);
  const restore = useRestoreRecycleBin(containerId, driveId);
  const purge = usePurgeRecycleBin(containerId);
  const [selected, setSelected] = useState<Set<string>>(new Set());

  const items = bin.data?.value ?? [];
  const selectedIds = useMemo(() => Array.from(selected), [selected]);

  const toggle = (id: string) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const restoreSelected = async () => {
    if (selectedIds.length === 0) return;
    try {
      await restore.mutateAsync(selectedIds);
      notify({ title: `Restored ${selectedIds.length} item${selectedIds.length === 1 ? "" : "s"}`, intent: "success" });
      setSelected(new Set());
    } catch (error) {
      notify({ title: "Restore failed", body: error instanceof Error ? error.message : "Unknown error", intent: "error" });
    }
  };

  const purgeSelected = async () => {
    if (selectedIds.length === 0) return;
    try {
      await purge.mutateAsync(selectedIds);
      notify({ title: `Permanently deleted ${selectedIds.length} item${selectedIds.length === 1 ? "" : "s"}`, intent: "success" });
      setSelected(new Set());
    } catch (error) {
      notify({ title: "Permanent delete failed", body: error instanceof Error ? error.message : "Unknown error", intent: "error" });
    }
  };

  return (
    <Drawer open={open} onOpenChange={(_, data) => onOpenChange(data.open)} position="end" size="medium">
      <DrawerHeader>
        <DrawerHeaderTitle action={<Button appearance="subtle" icon={<DismissRegular />} onClick={() => onOpenChange(false)} aria-label="Close" />}>
          Recycle bin
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody className={styles.body}>
        <Toolbar className={styles.toolbar}>
          <ToolbarButton appearance="primary" icon={<ArrowClockwiseRegular />} disabled={selectedIds.length === 0 || restore.isPending} onClick={() => void restoreSelected()}>
            Restore ({selectedIds.length})
          </ToolbarButton>
          <ToolbarButton icon={<DeleteRegular />} disabled={selectedIds.length === 0 || purge.isPending} onClick={() => void purgeSelected()}>
            Delete forever
          </ToolbarButton>
          <ToolbarButton onClick={() => void bin.refetch()} disabled={bin.isFetching}>
            Refresh
          </ToolbarButton>
        </Toolbar>
        {bin.isLoading ? <Spinner label="Loading deleted items" /> : items.length === 0 ? (
          <div className={styles.empty}>
            <Text>Nothing in the recycle bin.</Text>
          </div>
        ) : items.map((item) => (
          <div key={item.id} className={styles.row}>
            <Checkbox checked={selected.has(item.id)} onChange={() => toggle(item.id)} aria-label={`Select ${item.name ?? item.id}`} />
            <div className={styles.meta}>
              <Text weight="semibold">{item.name ?? "Untitled"}</Text>
              <Text size={200}>{formatBytes(item.size ?? 0)} • Deleted {item.deletedDateTime ? formatDate(item.deletedDateTime) : "—"}{item.deletedBy?.displayName ? ` by ${item.deletedBy.displayName}` : ""}</Text>
            </div>
            <Button size="small" appearance="subtle" onClick={() => { setSelected(new Set([item.id])); void restoreSelected(); }} disabled={restore.isPending}>
              Restore
            </Button>
          </div>
        ))}
      </DrawerBody>
    </Drawer>
  );
}
