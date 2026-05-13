"use client";

import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, MessageBar, MessageBarBody, MessageBarTitle, Text } from "@fluentui/react-components";

export function DeleteConfirmDialog({ itemName, open, onOpenChange, onDelete }: { itemName?: string; open: boolean; onOpenChange: (open: boolean) => void; onDelete: () => Promise<void> | void }) {
  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Delete {itemName}</DialogTitle>
          <DialogContent>
            <MessageBar intent="warning">
              <MessageBarBody>
                <MessageBarTitle>This permanently removes the item.</MessageBarTitle>
                Shared links and invitations to this item will stop working after deletion.
              </MessageBarBody>
            </MessageBar>
            <Text block>Choose “Permanently delete” to continue.</Text>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => onOpenChange(false)}>Cancel</Button>
            <Button appearance="primary" onClick={() => void onDelete()}>Permanently delete</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
