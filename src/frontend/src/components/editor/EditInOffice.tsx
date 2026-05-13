"use client";

import { Button, Link, Text, makeStyles, tokens } from "@fluentui/react-components";
import { EditRegular } from "@fluentui/react-icons";
import { useState } from "react";
import type { DriveItem } from "@/lib/api/types";
import { officeAppName } from "@/lib/files";
import { useAppToast } from "@/components/common/Toaster";

const useStyles = makeStyles({
  fallback: {
    display: "grid",
    gap: tokens.spacingVerticalXS,
  },
});

export function EditInOffice({ item, onOpened }: { item: DriveItem; onOpened?: () => void }) {
  const styles = useStyles();
  const notify = useAppToast();
  const [blocked, setBlocked] = useState(false);
  const appName = officeAppName(item);

  const openEditor = () => {
    const child = window.open(item.webUrl, "_blank", "noopener,noreferrer");
    if (child === null) {
      setBlocked(true);
      notify({ title: "Popup blocked", body: "Use the direct link to open Office in a new tab.", intent: "warning" });
      return;
    }
    onOpened?.();
  };

  return (
    <div className={styles.fallback}>
      <Button appearance="primary" icon={<EditRegular />} onClick={openEditor}>Edit in {appName}</Button>
      {blocked ? (
        <Text size={200}>
          Popup blocked. <Link href={item.webUrl} target="_blank" rel="noopener noreferrer">Open in {appName}</Link>
        </Text>
      ) : null}
    </div>
  );
}
