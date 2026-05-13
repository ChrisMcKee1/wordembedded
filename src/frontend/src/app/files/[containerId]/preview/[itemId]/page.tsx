"use client";

import { Accordion, AccordionHeader, AccordionItem, AccordionPanel, Button, Card, Divider, Link, Spinner, Switch, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowLeftRegular, ShareRegular } from "@fluentui/react-icons";
import { useParams, useRouter, useSearchParams } from "next/navigation";
import { useEffect, useState } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { AuthBoundary } from "@/components/auth/AuthBoundary";
import { AppShell } from "@/components/layout/AppShell";
import { ShareDialog } from "@/components/dialogs/ShareDialog";
import { VersionsDrawer } from "@/components/dialogs/VersionsDrawer";
import { EditInOffice } from "@/components/editor/EditInOffice";
import { useContainer } from "@/lib/hooks/use-containers";
import { useItem } from "@/lib/hooks/use-item";
import { usePreview } from "@/lib/hooks/use-preview";
import { WorkflowControls } from "@/components/workflow/WorkflowControls";
import { extensionFor, formatBytes, formatDate, modifiedByName, previewUrlWithNoBanner } from "@/lib/files";
import { childrenKey } from "@/lib/hooks/use-children";
import type { PreviewOptions } from "@/lib/api/preview";

const useStyles = makeStyles({
  stack: { display: "grid", gap: tokens.spacingVerticalL },
  top: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: tokens.spacingHorizontalM, flexWrap: "wrap" },
  layout: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) 340px", gap: tokens.spacingHorizontalL, "@media (max-width: 980px)": { gridTemplateColumns: "1fr" } },
  preview: { minHeight: "70vh", padding: 0, overflow: "hidden" },
  iframe: { width: "100%", height: "70vh", backgroundColor: tokens.colorNeutralBackground2 },
  sidebar: { padding: tokens.spacingHorizontalL, display: "grid", gap: tokens.spacingVerticalM, alignContent: "start" },
  meta: { display: "grid", gap: tokens.spacingVerticalXS },
  actions: { display: "grid", gap: tokens.spacingVerticalS },
  toggleRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: tokens.spacingHorizontalS },
  pdfPanel: { display: "grid", gap: tokens.spacingVerticalXS, padding: tokens.spacingVerticalS, border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: tokens.borderRadiusMedium },
});

// Per Microsoft docs the SPE PDF viewer features (print + sticky notes + in-file search) are
// now baked into the preview iframe Graph returns from POST /drives/{id}/items/{id}/preview as
// of the March 2026 release. The older `?embed={"mpp":true,"mpsn":true}` query string on
// driveItem.webUrl is still documented for top-level navigation, but webUrl can't be iframed
// (CSP frame-ancestors) and direct nav requires an SPO session — so we don't expose it here.

export default function PreviewPage() {
  const styles = useStyles();
  const router = useRouter();
  const params = useParams<{ containerId: string; itemId: string }>();
  const searchParams = useSearchParams();
  const queryClient = useQueryClient();
  const containerId = decodeURIComponent(params.containerId);
  const itemId = decodeURIComponent(params.itemId);
  const container = useContainer(containerId);
  const driveId = searchParams.get("driveId") ?? container.data?.driveId;
  const item = useItem(driveId, itemId);
  const [shareOpen, setShareOpen] = useState(false);
  const [versionsOpen, setVersionsOpen] = useState(false);
  const [editMode, setEditMode] = useState(false);

  const isPdf = item.data ? (extensionFor(item.data.name) === "pdf" || (item.data.file?.mimeType?.includes("pdf") ?? false)) : false;

  const previewOptions: PreviewOptions | undefined = editMode
    ? { allowEdit: true, viewer: "office", chromeless: false }
    : undefined;
  const preview = usePreview(driveId, itemId, previewOptions);

  useEffect(() => {
    const onFocus = () => {
      if (!driveId) return;
      void queryClient.invalidateQueries({ queryKey: ["item", driveId, itemId] });
      void queryClient.invalidateQueries({ queryKey: childrenKey(driveId, item.data?.parentReference?.id) });
    };
    window.addEventListener("focus", onFocus);
    return () => window.removeEventListener("focus", onFocus);
  }, [driveId, itemId, item.data?.parentReference?.id, queryClient]);

  const rawPreviewUrl = preview.data?.getUrl;
  // Iframe always uses the embeddable preview getUrl. Per the March 2026 SPE release notes,
  // the native PDF viewer (in-file search, sticky notes, printing) is automatically wired into
  // the URL the preview API returns for PDFs — no extra query params needed.
  const previewUrl = editMode ? rawPreviewUrl : previewUrlWithNoBanner(rawPreviewUrl);

  return (
    <AuthBoundary>
      <AppShell>
        <section className={styles.stack}>
          <div className={styles.top}>
            <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={() => router.push(`/files/${encodeURIComponent(containerId)}`)}>Back to files</Button>
            <Text>{container.data?.displayName}</Text>
          </div>
          <div className={styles.layout}>
            <Card className={styles.preview} appearance="outline">
              {preview.isLoading ? <Spinner label="Loading preview" /> : previewUrl ? (
                <iframe className={styles.iframe} style={{ border: 0 }} src={previewUrl} title={item.data?.name ?? "Document preview"} onError={() => void preview.refetch()} allow="clipboard-read; clipboard-write" />
              ) : (
                <Text>Preview is unavailable for this file.</Text>
              )}
            </Card>
            <Card className={styles.sidebar} appearance="outline">
              {item.isLoading ? <Spinner label="Loading file details" /> : item.data ? (
                <>
                  <div className={styles.meta}>
                    <Text as="h1" size={600} weight="semibold">{item.data.name}</Text>
                    <Text>{formatBytes(item.data.size)}</Text>
                    <Text>Modified {formatDate(item.data.lastModifiedDateTime)}</Text>
                    <Text>Modified by {modifiedByName(item.data.lastModifiedBy)}</Text>
                  </div>
                  <Divider />
                  {isPdf ? (
                    <div className={styles.pdfPanel}>
                      <Text weight="semibold">SPE native PDF viewer</Text>
                      <Text size={200}>
                        The inline preview above is the SharePoint Embedded native PDF viewer. As of the March 2026 SPE update,
                        in-file <strong>search</strong>, <strong>sticky note</strong> rendering, and <strong>printing</strong>{" "}
                        are baked into the viewer Graph returns from <code>driveItem:preview</code> — no extra flags needed.
                      </Text>
                      <Text size={200}>
                        The <code>?embed=&#123;&quot;mpp&quot;:true,&quot;mpsn&quot;:true&#125;</code> override on <code>driveItem.webUrl</code>{" "}
                        is still documented for direct top-level navigation, but it requires an active SharePoint Online session for
                        the consuming tenant and can&apos;t be iframed (CSP <code>frame-ancestors</code>).
                      </Text>
                    </div>
                  ) : (
                    <>
                      <div className={styles.toggleRow}>
                        <Text>Edit inline (beta)</Text>
                        <Switch checked={editMode} onChange={(_, data) => setEditMode(data.checked)} />
                      </div>
                      <Text size={200}>Uses the Microsoft Graph beta <code>allowEdit</code> preview. If Office blocks the embed, click below to open it in a tab.</Text>
                    </>
                  )}
                  <div className={styles.actions}>
                    <EditInOffice item={item.data} onOpened={() => { if (driveId) void queryClient.invalidateQueries({ queryKey: ["item", driveId, itemId] }); }} />
                    <Button icon={<ShareRegular />} onClick={() => setShareOpen(true)}>Share</Button>
                  </div>
                  <Divider />
                  <WorkflowControls containerId={containerId} driveId={driveId} itemId={itemId} itemName={item.data.name} />
                  <Divider />
                  <Accordion collapsible defaultOpenItems={["versions"]}>
                    <AccordionItem value="versions">
                      <AccordionHeader onClick={() => setVersionsOpen(true)}>Versions</AccordionHeader>
                      <AccordionPanel><Text>Open the versions drawer to restore or download prior versions.</Text></AccordionPanel>
                    </AccordionItem>
                    <AccordionItem value="activity">
                      <AccordionHeader>Activity</AccordionHeader>
                      <AccordionPanel><Text>Activity will appear here when webhook history is connected.</Text></AccordionPanel>
                    </AccordionItem>
                  </Accordion>
                  <Link href={item.data.webUrl} target="_blank" rel="noopener noreferrer">Open Office link directly</Link>
                </>
              ) : <Text>File not found.</Text>}
            </Card>
          </div>
          <ShareDialog driveId={driveId} item={item.data} open={shareOpen} onOpenChange={setShareOpen} />
          <VersionsDrawer driveId={driveId} itemId={itemId} open={versionsOpen} onOpenChange={setVersionsOpen} />
        </section>
      </AppShell>
    </AuthBoundary>
  );
}
