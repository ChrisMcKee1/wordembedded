"use client";

import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Dropdown,
  Field,
  Input,
  Option,
  Tab,
  TabList,
  Text,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { CopyRegular } from "@fluentui/react-icons";
import { useState } from "react";
import type { DriveItem, LinkType, ShareRole } from "@/lib/api/types";
import { useCreateLink } from "@/lib/hooks/use-create-link";
import { useInvite } from "@/lib/hooks/use-invite";
import { useAppToast } from "@/components/common/Toaster";

const useStyles = makeStyles({
  form: {
    display: "grid",
    gap: tokens.spacingVerticalM,
  },
  row: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: tokens.spacingHorizontalM,
    "@media (max-width: 640px)": { gridTemplateColumns: "1fr" },
  },
  output: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    alignItems: "end",
  },
  grow: {
    flexGrow: 1,
  },
});

export function ShareDialog({ driveId, item, open, onOpenChange }: { driveId?: string; item?: DriveItem; open: boolean; onOpenChange: (open: boolean) => void }) {
  const styles = useStyles();
  const createLink = useCreateLink();
  const invite = useInvite();
  const notify = useAppToast();
  const [tab, setTab] = useState<"link" | "invite">("link");
  const [linkType, setLinkType] = useState<LinkType>("view");
  const [expiration, setExpiration] = useState("");
  const [password, setPassword] = useState("");
  const [linkUrl, setLinkUrl] = useState("");
  const [recipients, setRecipients] = useState("");
  const [role, setRole] = useState<ShareRole>("read");
  const [message, setMessage] = useState("");

  const createSharingLink = async () => {
    if (!driveId || !item) return;
    const response = await createLink.mutateAsync({
      driveId,
      itemId: item.id,
      request: {
        type: linkType,
        scope: "organization",
        password: password.trim() || undefined,
        expirationDateTime: expiration ? new Date(expiration).toISOString() : undefined,
      },
    });
    setLinkUrl(response.link.webUrl);
    notify({ title: "Sharing link created", intent: "success" });
  };

  const sendInvite = async () => {
    if (!driveId || !item) return;
    const people = recipients.split(/[;,\s]+/).map((email) => email.trim()).filter(Boolean).map((email) => ({ email }));
    if (people.length === 0) return;
    await invite.mutateAsync({ driveId, itemId: item.id, request: { recipients: people, roles: [role], sendInvitation: true, message: message.trim() || undefined } });
    notify({ title: "Invitation sent", intent: "success" });
    onOpenChange(false);
  };

  const copy = async () => {
    if (!linkUrl) return;
    await navigator.clipboard.writeText(linkUrl);
    notify({ title: "Link copied", intent: "success" });
  };

  return (
    <Dialog open={open} onOpenChange={(_, data) => onOpenChange(data.open)}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Share {item?.name}</DialogTitle>
          <DialogContent>
            <TabList selectedValue={tab} onTabSelect={(_, data) => setTab(data.value as "link" | "invite")}>
              <Tab value="link">Share link</Tab>
              <Tab value="invite">Invite people</Tab>
            </TabList>
            {tab === "link" ? (
              <div className={styles.form}>
                <Text>Phase 1 creates organization-scoped links only.</Text>
                <div className={styles.row}>
                  <Field label="Permission">
                    <Dropdown selectedOptions={[linkType]} value={linkType === "edit" ? "Edit" : "View"} onOptionSelect={(_, data) => setLinkType((data.optionValue as LinkType) ?? "view")}>
                      <Option value="view">View</Option>
                      <Option value="edit">Edit</Option>
                    </Dropdown>
                  </Field>
                  <Field label="Expiration">
                    <Input type="date" value={expiration} onChange={(_, data) => setExpiration(data.value)} />
                  </Field>
                </div>
                <Field label="Password (if tenant policy allows)">
                  <Input type="password" value={password} onChange={(_, data) => setPassword(data.value)} />
                </Field>
                {linkUrl ? (
                  <div className={styles.output}>
                    <Field label="Link" className={styles.grow}>
                      <Input readOnly value={linkUrl} />
                    </Field>
                    <Button icon={<CopyRegular />} onClick={() => void copy()}>Copy</Button>
                  </div>
                ) : null}
              </div>
            ) : (
              <div className={styles.form}>
                <Field label="People" hint="Separate email addresses with commas, semicolons, or spaces." required>
                  <Input value={recipients} onChange={(_, data) => setRecipients(data.value)} />
                </Field>
                <Field label="Role">
                  <Dropdown selectedOptions={[role]} value={role === "write" ? "Can edit" : "Can read"} onOptionSelect={(_, data) => setRole((data.optionValue as ShareRole) ?? "read")}>
                    <Option value="read">Can read</Option>
                    <Option value="write">Can edit</Option>
                  </Dropdown>
                </Field>
                <Field label="Message">
                  <Textarea value={message} onChange={(_, data) => setMessage(data.value)} />
                </Field>
              </div>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => onOpenChange(false)}>Close</Button>
            {tab === "link" ? <Button appearance="primary" onClick={() => void createSharingLink()} disabled={!driveId || !item || createLink.isPending}>Create link</Button> : <Button appearance="primary" onClick={() => void sendInvite()} disabled={!recipients.trim() || invite.isPending}>Send invite</Button>}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
