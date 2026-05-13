"use client";

import { Badge, Button, Link as FluentLink, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowResetRegular, ArrowSyncRegular, CheckmarkCircleRegular, CloudArrowUpRegular, EditRegular, OpenRegular, PersonClockRegular, SettingsRegular } from "@fluentui/react-icons";
import type { JSX } from "react";
import { useState } from "react";
import { useSetWorkflowStatus, useWorkflowStatus, usePublishToSharePoint, usePublishTarget } from "@/lib/hooks/use-workflow";
import { useAppToast } from "@/components/common/Toaster";
import { PublishTargetDialog } from "@/components/workflow/PublishTargetDialog";
import type { WorkflowStatus } from "@/lib/api/workflow";

const useStyles = makeStyles({
  bar: { display: "grid", gap: tokens.spacingVerticalS },
  row: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: tokens.spacingHorizontalS },
  actions: { display: "flex", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
});

const statusColor: Record<WorkflowStatus, "informative" | "warning" | "success" | "brand"> = {
  Draft: "informative",
  InReview: "warning",
  Approved: "success",
  Published: "brand",
};

const statusIcon: Record<WorkflowStatus, JSX.Element> = {
  Draft: <EditRegular />,
  InReview: <PersonClockRegular />,
  Approved: <CheckmarkCircleRegular />,
  Published: <CloudArrowUpRegular />,
};

const statusLabel: Record<WorkflowStatus, string> = {
  Draft: "Draft",
  InReview: "In review",
  Approved: "Approved",
  Published: "Published",
};

export function WorkflowControls({
  containerId,
  driveId,
  itemId,
  itemName,
}: {
  containerId: string;
  driveId?: string;
  itemId?: string;
  itemName?: string;
}) {
  const styles = useStyles();
  const notify = useAppToast();
  const status = useWorkflowStatus(driveId, itemId);
  const setStatus = useSetWorkflowStatus(driveId, itemId);
  const publish = usePublishToSharePoint(driveId);
  const target = usePublishTarget(containerId);
  const [targetDialogOpen, setTargetDialogOpen] = useState(false);

  if (!driveId || !itemId) return null;

  const current = (status.data?.status as WorkflowStatus | undefined) ?? "Draft";
  const published = status.data?.published;

  const transition = async (next: WorkflowStatus) => {
    try {
      await setStatus.mutateAsync(next);
      notify({ title: `Status: ${statusLabel[next]}`, intent: "success" });
    } catch (error) {
      notify({ title: "Could not update status", body: error instanceof Error ? error.message : "Unknown error", intent: "error" });
    }
  };

  const doPublish = async () => {
    if (!target.data?.driveId) {
      setTargetDialogOpen(true);
      return;
    }
    try {
      const result = await publish.mutateAsync({
        itemId: itemId!,
        targetDriveId: target.data.driveId,
        targetFolderId: target.data.folderId ?? undefined,
      });
      notify({
        title: `Published “${itemName ?? "file"}”`,
        body: result.monitorUrl ? "Graph accepted the copy to SharePoint." : "Published.",
        intent: "success",
      });
    } catch (error) {
      notify({ title: "Publish failed", body: error instanceof Error ? error.message : "Unknown error", intent: "error" });
    }
  };

  const busy = setStatus.isPending || publish.isPending;
  const hasTarget = Boolean(target.data?.driveId);

  return (
    <div className={styles.bar}>
      <div className={styles.row}>
        <Text weight="semibold">Workflow</Text>
        <Badge appearance="filled" color={statusColor[current]} icon={statusIcon[current]}>{statusLabel[current]}</Badge>
      </div>

      <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
        Publishing creates a <strong>copy</strong> in the destination library. The file you&apos;re editing here stays in this
        container; the published copy is a separate file. To overwrite the destination after further edits, click
        <strong> Publish to SharePoint</strong> again — it will create a new version of the destination file (same name).
      </Text>

      {current === "Published" && published?.webUrl ? (
        <div className={styles.row}>
          <Text size={200}>
            Published as <code>{published.name ?? "destination file"}</code>
            {published.publishedAt ? <> on <time dateTime={published.publishedAt}>{new Date(published.publishedAt).toLocaleString()}</time></> : null}
          </Text>
          <FluentLink href={published.webUrl} target="_blank" rel="noopener noreferrer">
            <OpenRegular /> View published version
          </FluentLink>
        </div>
      ) : null}

      <div className={styles.actions}>
        {current === "Draft" ? (
          <Button icon={<PersonClockRegular />} disabled={busy} onClick={() => void transition("InReview")}>Submit for review</Button>
        ) : null}
        {current === "InReview" ? (
          <>
            <Button appearance="primary" icon={<CheckmarkCircleRegular />} disabled={busy} onClick={() => void transition("Approved")}>Approve</Button>
            <Button icon={<ArrowResetRegular />} disabled={busy} onClick={() => void transition("Draft")}>Return to draft</Button>
          </>
        ) : null}
        {current === "Approved" ? (
          <>
            <Button appearance="primary" icon={<CloudArrowUpRegular />} disabled={busy || !hasTarget} onClick={() => void doPublish()}>Publish to SharePoint</Button>
            <Button icon={<ArrowResetRegular />} disabled={busy} onClick={() => void transition("InReview")}>Back to review</Button>
          </>
        ) : null}
        {current === "Published" ? (
          <>
            <Button appearance="primary" icon={<CloudArrowUpRegular />} disabled={busy || !hasTarget} onClick={() => void doPublish()}>Republish (overwrite)</Button>
            <Button icon={<ArrowSyncRegular />} disabled={busy} onClick={() => void transition("Draft")}>Reset to draft</Button>
          </>
        ) : null}
      </div>

      <div className={styles.row}>
        <Text size={200}>
          {hasTarget
            ? `Target: ${target.data?.folderPath || "(library root)"}`
            : "No publish target configured for this container."}
        </Text>
        <Button size="small" appearance="subtle" icon={<SettingsRegular />} onClick={() => setTargetDialogOpen(true)}>
          {hasTarget ? "Change" : "Configure"}
        </Button>
      </div>

      <PublishTargetDialog containerId={containerId} open={targetDialogOpen} onOpenChange={setTargetDialogOpen} />
    </div>
  );
}
