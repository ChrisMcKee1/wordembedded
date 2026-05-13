"use client";

import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Spinner, Text, makeStyles, tokens } from "@fluentui/react-components";
import { useEffect, useState } from "react";
import { useAppToast } from "@/components/common/Toaster";
import { useSetPublishTarget, useSharePointDrives, useSharePointSites, usePublishTarget } from "@/lib/hooks/use-workflow";

const useStyles = makeStyles({
  body: { display: "grid", gap: tokens.spacingVerticalM },
  pickerList: { display: "grid", gap: tokens.spacingVerticalXXS, maxHeight: "220px", overflowY: "auto", border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: tokens.borderRadiusMedium, padding: tokens.spacingVerticalXXS },
  pickerRow: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: tokens.spacingVerticalS, borderRadius: tokens.borderRadiusMedium, cursor: "pointer", gap: tokens.spacingHorizontalS },
  pickerRowActive: { backgroundColor: tokens.colorBrandBackground2 },
  searchRow: { display: "flex", gap: tokens.spacingHorizontalS, alignItems: "center" },
  empty: { padding: tokens.spacingVerticalS, color: tokens.colorNeutralForeground2 },
});

export function PublishTargetDialog({
  containerId,
  open,
  onOpenChange,
}: {
  containerId: string;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}) {
  const styles = useStyles();
  const notify = useAppToast();
  const current = usePublishTarget(containerId);
  const setTarget = useSetPublishTarget(containerId);
  const [siteQuery, setSiteQuery] = useState("");
  const [siteId, setSiteId] = useState<string | undefined>();
  const [driveId, setDriveId] = useState<string | undefined>();
  const [folderPath, setFolderPath] = useState("");

  const sites = useSharePointSites(siteQuery || undefined);
  const drives = useSharePointDrives(siteId);

  useEffect(() => {
    if (open && current.data) {
      setSiteId(current.data.siteId ?? undefined);
      setDriveId(current.data.driveId ?? undefined);
      setFolderPath(current.data.folderPath ?? "");
    }
  }, [open, current.data]);

  const save = async () => {
    if (!driveId) {
      notify({ title: "Pick a library", body: "Select a SharePoint site and document library first.", intent: "warning" });
      return;
    }
    try {
      await setTarget.mutateAsync({
        siteId: siteId ?? null,
        driveId,
        folderId: null,
        folderPath: folderPath.trim() || null,
      });
      notify({ title: "Publish target saved", intent: "success" });
      onOpenChange(false);
    } catch (error) {
      notify({ title: "Could not save target", body: error instanceof Error ? error.message : "Unknown error", intent: "error" });
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Container publish target</DialogTitle>
          <DialogContent className={styles.body}>
            <Text size={200}>When an approved file is published, the app copies it to this SharePoint document library.</Text>

            <Field label="SharePoint site">
              <div className={styles.searchRow}>
                <Input
                  value={siteQuery}
                  onChange={(_, data) => setSiteQuery(data.value)}
                  placeholder="Filter sites…"
                  style={{ flex: 1 }}
                />
              </div>
              <div className={styles.pickerList} role="listbox" aria-label="SharePoint sites">
                {sites.isLoading ? <Spinner size="tiny" label="Loading sites" /> : null}
                {sites.data && sites.data.value.length === 0 ? <div className={styles.empty}>No sites found.</div> : null}
                {sites.data?.value.map((site) => {
                  const active = site.id === siteId;
                  return (
                    <div
                      key={site.id}
                      role="option"
                      aria-selected={active}
                      tabIndex={0}
                      className={`${styles.pickerRow} ${active ? styles.pickerRowActive : ""}`}
                      onClick={() => { setSiteId(site.id); setDriveId(undefined); }}
                      onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { setSiteId(site.id); setDriveId(undefined); } }}
                    >
                      <div style={{ minWidth: 0 }}>
                        <Text weight="semibold">{site.displayName ?? site.name ?? site.id}</Text>
                        <Text block size={100} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{site.webUrl}</Text>
                      </div>
                    </div>
                  );
                })}
              </div>
            </Field>

            <Field label="Document library">
              {!siteId ? <Text size={200} className={styles.empty}>Pick a site first.</Text> : (
                <div className={styles.pickerList} role="listbox" aria-label="Document libraries">
                  {drives.isLoading ? <Spinner size="tiny" label="Loading libraries" /> : null}
                  {drives.data && drives.data.value.length === 0 ? <div className={styles.empty}>No document libraries on this site.</div> : null}
                  {drives.data?.value.map((drive) => {
                    const active = drive.id === driveId;
                    return (
                      <div
                        key={drive.id}
                        role="option"
                        aria-selected={active}
                        tabIndex={0}
                        className={`${styles.pickerRow} ${active ? styles.pickerRowActive : ""}`}
                        onClick={() => setDriveId(drive.id)}
                        onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") setDriveId(drive.id); }}
                      >
                        <Text weight="semibold">{drive.name ?? drive.id}</Text>
                        <Text size={200}>{drive.driveType ?? ""}</Text>
                      </div>
                    );
                  })}
                </div>
              )}
            </Field>

            <Field label="Folder path (optional)" hint="Leave blank to publish to the library root.">
              <Input value={folderPath} onChange={(_, data) => setFolderPath(data.value)} placeholder="e.g. Published/2026" />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => onOpenChange(false)}>Cancel</Button>
            <Button appearance="primary" disabled={setTarget.isPending || !driveId} onClick={() => void save()}>Save target</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
