"use client";

import { makeStyles, shorthands, tokens } from "@fluentui/react-components";
import type { ReactNode } from "react";
import { useState } from "react";

const useStyles = makeStyles({
  root: {
    position: "relative",
  },
  overlay: {
    position: "absolute",
    top: 0,
    right: 0,
    bottom: 0,
    left: 0,
    zIndex: 1,
    display: "grid",
    placeItems: "center",
    backgroundColor: tokens.colorNeutralBackground1Hover,
    ...shorthands.border(tokens.strokeWidthThick, "dashed", tokens.colorBrandStroke1),
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
  },
});

interface UploadZoneProps {
  children: ReactNode;
  onFiles: (files: File[]) => void;
  disabled?: boolean;
}

export function UploadZone({ children, disabled, onFiles }: UploadZoneProps) {
  const styles = useStyles();
  const [isDragging, setIsDragging] = useState(false);

  return (
    <div
      className={styles.root}
      onDragEnter={(event) => {
        if (disabled) return;
        event.preventDefault();
        setIsDragging(true);
      }}
      onDragOver={(event) => {
        if (disabled) return;
        event.preventDefault();
      }}
      onDragLeave={(event) => {
        if (event.currentTarget.contains(event.relatedTarget as Node | null)) return;
        setIsDragging(false);
      }}
      onDrop={(event) => {
        if (disabled) return;
        event.preventDefault();
        setIsDragging(false);
        onFiles(Array.from(event.dataTransfer.files));
      }}
    >
      {isDragging ? <div className={styles.overlay}>Drop files to upload</div> : null}
      {children}
    </div>
  );
}
