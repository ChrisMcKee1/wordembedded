"use client";

import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input } from "@fluentui/react-components";
import { useState } from "react";

export function RenameDialog({ currentName, open, onOpenChange, onRename }: { currentName?: string; open: boolean; onOpenChange: (open: boolean) => void; onRename: (name: string) => Promise<void> | void }) {
  const [name, setName] = useState(currentName ?? "");
  const [isSaving, setIsSaving] = useState(false);

  const rename = async () => {
    const trimmed = name.trim();
    if (!trimmed || trimmed === currentName) return;
    setIsSaving(true);
    try {
      await onRename(trimmed);
      onOpenChange(false);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Rename</DialogTitle>
          <DialogContent>
            <Field label="Name" required>
              <Input value={name} onChange={(_, data) => setName(data.value)} onKeyDown={(event) => { if (event.key === "Enter") void rename(); }} autoFocus />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => onOpenChange(false)}>Cancel</Button>
            <Button appearance="primary" disabled={!name.trim() || name.trim() === currentName || isSaving} onClick={() => void rename()}>Save</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
