"use client";

import { Button, DrawerBody, DrawerHeader, DrawerHeaderTitle, Link, OverlayDrawer, Text, makeStyles, shorthands, tokens } from "@fluentui/react-components";
import { DismissRegular } from "@fluentui/react-icons";
import { useVersions } from "@/lib/hooks/use-versions";
import { useRestoreVersion } from "@/lib/hooks/use-restore-version";
import { formatBytes, formatDate, modifiedByName } from "@/lib/files";

const useStyles = makeStyles({
  list: {
    display: "grid",
    gap: tokens.spacingVerticalM,
  },
  row: {
    display: "grid",
    gap: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalM,
    ...shorthands.borderBottom(tokens.strokeWidthThin, "solid", tokens.colorNeutralStroke2),
  },
  actions: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
  },
});

export function VersionsDrawer({ driveId, itemId, open, onOpenChange }: { driveId?: string; itemId?: string; open: boolean; onOpenChange: (open: boolean) => void }) {
  const styles = useStyles();
  const versions = useVersions(driveId, itemId);
  const restore = useRestoreVersion();

  return (
    <OverlayDrawer open={open} onOpenChange={(_, data) => onOpenChange(data.open)} position="end">
      <DrawerHeader>
        <DrawerHeaderTitle action={<Button appearance="subtle" aria-label="Close versions" icon={<DismissRegular />} onClick={() => onOpenChange(false)} />}>Versions</DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody>
        <div className={styles.list}>
          {versions.data?.value.map((version) => (
            <div key={version.id} className={styles.row}>
              <Text weight="semibold">Version {version.id}</Text>
              <Text>{formatDate(version.lastModifiedDateTime)} · {formatBytes(version.size)}</Text>
              <Text size={200}>Modified by {modifiedByName(version.lastModifiedBy)}</Text>
              <div className={styles.actions}>
                <Button size="small" onClick={() => driveId && itemId && restore.mutate({ driveId, itemId, versionId: version.id })}>Restore</Button>
                {version.downloadUrl ? <Link href={version.downloadUrl}>Download</Link> : null}
              </div>
            </div>
          )) ?? <Text>No versions found.</Text>}
        </div>
      </DrawerBody>
    </OverlayDrawer>
  );
}
