"use client";

import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Textarea } from "@fluentui/react-components";
import { useState } from "react";

export interface NewContainerSubmit {
  displayName: string;
  description?: string;
}

export function NewContainerDialog({
  open,
  onOpenChange,
  onCreate,
}: {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onCreate: (input: NewContainerSubmit) => Promise<void> | void;
}) {
  const [displayName, setDisplayName] = useState("");
  const [description, setDescription] = useState("");
  const [isSaving, setIsSaving] = useState(false);

  const create = async () => {
    const trimmedName = displayName.trim();
    if (!trimmedName) return;
    setIsSaving(true);
    try {
      const trimmedDescription = description.trim();
      await onCreate({ displayName: trimmedName, description: trimmedDescription || undefined });
      setDisplayName("");
      setDescription("");
      onOpenChange(false);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>New container</DialogTitle>
          <DialogContent>
            <Field label="Display name" required hint="Shows up wherever this workspace is listed.">
              <Input
                value={displayName}
                onChange={(_, data) => setDisplayName(data.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter" && !event.shiftKey) {
                    event.preventDefault();
                    void create();
                  }
                }}
                autoFocus
                maxLength={150}
              />
            </Field>
            <Field label="Description (optional)">
              <Textarea
                value={description}
                onChange={(_, data) => setDescription(data.value)}
                resize="vertical"
                rows={3}
                maxLength={1024}
              />
            </Field>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => onOpenChange(false)}>Cancel</Button>
            <Button appearance="primary" disabled={!displayName.trim() || isSaving} onClick={() => void create()}>
              {isSaving ? "Creating…" : "Create"}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
