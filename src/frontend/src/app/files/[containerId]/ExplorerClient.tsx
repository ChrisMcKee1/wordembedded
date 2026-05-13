"use client";

import { Button, Card, Spinner, Text, ToggleButton, makeStyles, tokens, type TableRowId } from "@fluentui/react-components";
import { ArrowUploadRegular, CloudArrowUpRegular, DeleteRegular, DeleteRegular as RecycleIcon, FolderAddRegular, GridRegular, ListRegular } from "@fluentui/react-icons";
import { useParams } from "next/navigation";
import { useRef, useState } from "react";
import { AuthBoundary } from "@/components/auth/AuthBoundary";
import { AppShell } from "@/components/layout/AppShell";
import { FileBreadcrumb } from "@/components/files/Breadcrumb";
import { FileList } from "@/components/files/FileList";
import { UploadZone } from "@/components/files/UploadZone";
import { UploadProgress, type UploadProgressItem } from "@/components/files/UploadProgress";
import { NewFolderDialog } from "@/components/dialogs/NewFolderDialog";
import { RenameDialog } from "@/components/dialogs/RenameDialog";
import { DeleteConfirmDialog } from "@/components/dialogs/DeleteConfirmDialog";
import { ShareDialog } from "@/components/dialogs/ShareDialog";
import { VersionsDrawer } from "@/components/dialogs/VersionsDrawer";
import { RecycleBinDrawer } from "@/components/dialogs/RecycleBinDrawer";
import { SharePointSearchPanel } from "@/components/sharepoint/SharePointSearchPanel";
import { PublishTargetDialog } from "@/components/workflow/PublishTargetDialog";
import { ThumbnailTile } from "@/components/files/ThumbnailTile";
import { useAppToast } from "@/components/common/Toaster";
import type { DriveItem } from "@/lib/api/types";
import { useContainer } from "@/lib/hooks/use-containers";
import { useChildren } from "@/lib/hooks/use-children";
import { useCreateFolder } from "@/lib/hooks/use-create-folder";
import { useDelete } from "@/lib/hooks/use-delete";
import { useRename } from "@/lib/hooks/use-rename";
import { useUpload } from "@/lib/hooks/use-upload";
import { decodeRouteSegment, formatBytes, formatDate, type RoutePathSegment } from "@/lib/files";

const useStyles = makeStyles({
  stack: { display: "grid", gap: tokens.spacingVerticalL },
  header: { display: "grid", gap: tokens.spacingVerticalS },
  bar: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: tokens.spacingHorizontalM, flexWrap: "wrap" },
  actions: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
  panel: { padding: tokens.spacingHorizontalM },
  grid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))", gap: tokens.spacingHorizontalM },
  tile: { padding: tokens.spacingHorizontalM, display: "grid", gap: tokens.spacingVerticalS, cursor: "pointer" },
});

function ExplorerContent({ containerId, pathSegments }: { containerId: string; pathSegments: RoutePathSegment[] }) {
  const styles = useStyles();
  const notify = useAppToast();
  const inputRef = useRef<HTMLInputElement>(null);
  const [selectedItems, setSelectedItems] = useState<Set<TableRowId>>(new Set());
  const [view, setView] = useState<"list" | "grid">("list");
  const [dialog, setDialog] = useState<"new" | "rename" | "delete" | "share" | undefined>();
  const [activeItem, setActiveItem] = useState<DriveItem | undefined>();
  const [versionsItem, setVersionsItem] = useState<DriveItem | undefined>();
  const [uploads, setUploads] = useState<UploadProgressItem[]>([]);
  const [recycleOpen, setRecycleOpen] = useState(false);
  const [publishTargetOpen, setPublishTargetOpen] = useState(false);
  const container = useContainer(containerId);
  const driveId = container.data?.driveId;
  const parentId = pathSegments.at(-1)?.id ?? "root";
  const currentFolderId = pathSegments.at(-1)?.id;
  const children = useChildren(driveId, currentFolderId);
  const createFolder = useCreateFolder();
  const rename = useRename();
  const deleteMutation = useDelete();
  const upload = useUpload();

  const uploadFiles = (files: File[]) => {
    if (!driveId || files.length === 0) return;
    for (const file of files) {
      const id = `${file.name}-${Date.now()}-${Math.random()}`;
      setUploads((current) => [...current, { id, name: file.name, progress: 0, status: "uploading" }]);
      upload.mutate(
        { driveId, parentId, file, onProgress: (loaded, total) => setUploads((current) => current.map((item) => item.id === id ? { ...item, progress: total ? (loaded / total) * 100 : 0 } : item)) },
        {
          onSuccess: () => {
            setUploads((current) => current.map((item) => item.id === id ? { ...item, progress: 100, status: "done" } : item));
            notify({ title: "Upload complete", body: file.name, intent: "success" });
          },
          onError: (error) => {
            setUploads((current) => current.map((item) => item.id === id ? { ...item, status: "error" } : item));
            notify({ title: "Upload failed", body: error instanceof Error ? error.message : file.name, intent: "error" });
          },
        },
      );
    }
  };

  const bulkDelete = async () => {
    if (!driveId || selectedItems.size === 0) return;
    await Promise.all(Array.from(selectedItems).map((id) => deleteMutation.mutateAsync({ driveId, itemId: String(id), parentId: currentFolderId })));
    setSelectedItems(new Set());
    notify({ title: "Selected items deleted", intent: "success" });
  };

  const items = children.data?.value ?? [];

  return (
    <section className={styles.stack}>
      <div className={styles.header}>
        <FileBreadcrumb containerId={containerId} containerName={container.data?.displayName} segments={pathSegments} />
        <div className={styles.bar}>
          <div>
            <Text as="h1" size={800} weight="semibold">{pathSegments.at(-1)?.name ?? container.data?.displayName ?? "Files"}</Text>
            <Text block>{items.length} item{items.length === 1 ? "" : "s"}</Text>
          </div>
          <div className={styles.actions}>
            {selectedItems.size > 0 ? <Button icon={<DeleteRegular />} onClick={() => void bulkDelete()}>Delete selected</Button> : null}
            <Button icon={<ArrowUploadRegular />} onClick={() => inputRef.current?.click()}>Upload</Button>
            <input ref={inputRef} type="file" multiple hidden onChange={(event) => uploadFiles(Array.from(event.currentTarget.files ?? []))} />
            <Button icon={<FolderAddRegular />} onClick={() => setDialog("new")}>New folder</Button>
            <Button icon={<RecycleIcon />} onClick={() => setRecycleOpen(true)}>Recycle bin</Button>
            <Button icon={<CloudArrowUpRegular />} onClick={() => setPublishTargetOpen(true)}>Publish target</Button>
            <ToggleButton checked={view === "list"} icon={<ListRegular />} onClick={() => setView("list")}>List</ToggleButton>
            <ToggleButton checked={view === "grid"} icon={<GridRegular />} onClick={() => setView("grid")}>Grid</ToggleButton>
          </div>
        </div>
      </div>
      <UploadProgress uploads={uploads} />
      <SharePointSearchPanel targetDriveId={driveId} targetFolderId={currentFolderId} />
      <UploadZone onFiles={uploadFiles} disabled={!driveId}>
        <Card className={styles.panel} appearance="outline">
          {children.isLoading || container.isLoading ? <Spinner label="Loading files" /> : view === "list" ? (
            <FileList containerId={containerId} driveId={driveId ?? ""} items={items} pathSegments={pathSegments} selectedItems={selectedItems} onSelectionChange={setSelectedItems} onRename={(item) => { setActiveItem(item); setDialog("rename"); }} onDelete={(item) => { setActiveItem(item); setDialog("delete"); }} onShare={(item) => { setActiveItem(item); setDialog("share"); }} onVersions={setVersionsItem} />
          ) : (
            <div className={styles.grid}>{items.map((item) => <Card key={item.id} className={styles.tile} appearance="subtle"><ThumbnailTile driveId={driveId} item={item} /><Text weight="semibold">{item.name}</Text><Text size={200}>{item.folder ? `${item.folder.childCount} items` : formatBytes(item.size)}</Text><Text size={200}>Modified {formatDate(item.lastModifiedDateTime)}</Text></Card>)}</div>
          )}
        </Card>
      </UploadZone>
      <NewFolderDialog open={dialog === "new"} onOpenChange={(open) => setDialog(open ? "new" : undefined)} onCreate={async (name) => { if (driveId) { await createFolder.mutateAsync({ driveId, parentId, name }); notify({ title: "Folder created", intent: "success" }); } }} />
      <RenameDialog key={activeItem?.id ?? "rename"} currentName={activeItem?.name} open={dialog === "rename"} onOpenChange={(open) => setDialog(open ? "rename" : undefined)} onRename={async (name) => { if (driveId && activeItem) { await rename.mutateAsync({ driveId, itemId: activeItem.id, parentId: currentFolderId, name }); notify({ title: "Item renamed", intent: "success" }); } }} />
      <DeleteConfirmDialog itemName={activeItem?.name} open={dialog === "delete"} onOpenChange={(open) => setDialog(open ? "delete" : undefined)} onDelete={async () => { if (driveId && activeItem) { await deleteMutation.mutateAsync({ driveId, itemId: activeItem.id, parentId: currentFolderId }); setDialog(undefined); notify({ title: "Item deleted", intent: "success" }); } }} />
      <ShareDialog driveId={driveId} item={activeItem} open={dialog === "share"} onOpenChange={(open) => setDialog(open ? "share" : undefined)} />
      <VersionsDrawer driveId={driveId} itemId={versionsItem?.id} open={Boolean(versionsItem)} onOpenChange={(open) => { if (!open) setVersionsItem(undefined); }} />
      <RecycleBinDrawer containerId={containerId} driveId={driveId} open={recycleOpen} onOpenChange={setRecycleOpen} />
      <PublishTargetDialog containerId={containerId} open={publishTargetOpen} onOpenChange={setPublishTargetOpen} />
    </section>
  );
}

export function ExplorerPage() {
  const params = useParams<{ containerId: string; path?: string[] }>();
  const containerId = decodeURIComponent(params.containerId);
  const pathSegments = (params.path ?? []).map(decodeRouteSegment);
  return (
    <AuthBoundary>
      <AppShell>
        <ExplorerContent containerId={containerId} pathSegments={pathSegments} />
      </AppShell>
    </AuthBoundary>
  );
}
