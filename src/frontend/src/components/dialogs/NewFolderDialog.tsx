"use client";

import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input } from "@fluentui/react-components";
import { useState } from "react";

export function NewFolderDialog({ open, onOpenChange, onCreate }: { open: boolean; onOpenChange: (open: boolean) => void; onCreate: (name: string) => Promise<void> | void }) {
  const [name, setName] = useState("");
  const [isSaving, setIsSaving] = useState(false);

  const create = async () => {
    const trimmed = name.trim();
    if (!trimmed) return;
    setIsSaving(true);
    try {
      await onCreate(trimmed);
      setName("");
      onOpenChange(false);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>New folder</DialogTitle>
          <DialogContent>
            <Field label="Folder name" required>
              <Input value={name} onChange={(_, data) => setName(data.value)} onKeyDown={(event) => { if (event.key === "Enter") void create(); }} autoFocus />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => onOpenChange(false)}>Cancel</Button>
            <Button appearance="primary" disabled={!name.trim() || isSaving} onClick={() => void create()}>Create</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
