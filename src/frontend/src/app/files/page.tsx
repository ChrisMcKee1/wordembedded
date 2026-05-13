"use client";

import { Badge, Button, Card, Spinner, Text, makeStyles, tokens } from "@fluentui/react-components";
import { AddRegular, FolderRegular } from "@fluentui/react-icons";
import { useRouter } from "next/navigation";
import { useEffect, useState } from "react";
import { AuthBoundary } from "@/components/auth/AuthBoundary";
import { AppShell } from "@/components/layout/AppShell";
import { useAppToast } from "@/components/common/Toaster";
import { useContainers, useCreateContainer } from "@/lib/hooks/use-containers";
import { useChildren } from "@/lib/hooks/use-children";
import type { Container } from "@/lib/api/types";
import { formatDate } from "@/lib/files";
import { NewContainerDialog, type NewContainerSubmit } from "@/components/dialogs/NewContainerDialog";

const useStyles = makeStyles({
  stack: { display: "grid", gap: tokens.spacingVerticalL },
  header: { display: "flex", justifyContent: "space-between", alignItems: "start", gap: tokens.spacingHorizontalL, flexWrap: "wrap" },
  cards: { display: "grid", gridTemplateColumns: "minmax(260px, 1.3fr) minmax(260px, 1fr)", gap: tokens.spacingHorizontalL, "@media (max-width: 860px)": { gridTemplateColumns: "1fr" } },
  sideCards: { display: "grid", gap: tokens.spacingVerticalM },
  card: { padding: tokens.spacingHorizontalL, cursor: "pointer" },
  featured: { minHeight: "220px", display: "grid", alignContent: "space-between" },
  cardHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: tokens.spacingHorizontalM },
  preview: { display: "grid", gap: tokens.spacingVerticalXS, marginTop: tokens.spacingVerticalM, color: tokens.colorNeutralForeground2 },
  empty: { minHeight: "320px", display: "grid", placeItems: "center", textAlign: "center", gap: tokens.spacingVerticalM },
});

function ContainerCard({ container, featured }: { container: Container; featured?: boolean }) {
  const styles = useStyles();
  const router = useRouter();
  const children = useChildren(container.driveId);
  const recent = children.data?.value.slice(0, featured ? 4 : 2) ?? [];

  return (
    <Card className={`${styles.card} ${featured ? styles.featured : ""}`} appearance="outline" onClick={() => router.push(`/files/${encodeURIComponent(container.id)}`)}>
      <div className={styles.cardHeader}>
        <div>
          <Text weight="semibold" size={featured ? 600 : 500}>{container.displayName}</Text>
          <Text block size={200}>Created {formatDate(container.createdDateTime)}</Text>
        </div>
        {featured ? <Badge appearance="filled" color="brand">Pinned</Badge> : <Badge>{container.status}</Badge>}
      </div>
      <div className={styles.preview}>
        <Text weight="semibold" size={200}>Recent at root</Text>
        {children.isLoading ? <Text size={200}>Loading files…</Text> : recent.length > 0 ? recent.map((item) => <Text key={item.id} size={200}>• {item.name}</Text>) : <Text size={200}>No files at root</Text>}
      </div>
    </Card>
  );
}

export default function FilesPage() {
  const styles = useStyles();
  const router = useRouter();
  const notify = useAppToast();
  const containers = useContainers();
  const create = useCreateContainer();
  const [dialogOpen, setDialogOpen] = useState(false);

  useEffect(() => {
    if (containers.data?.length === 1) router.push(`/files/${encodeURIComponent(containers.data[0].id)}`);
  }, [containers.data, router]);

  const handleCreate = async (input: NewContainerSubmit) => {
    try {
      const created = await create.mutateAsync(input);
      notify({ title: `Created “${created.displayName}”`, intent: "success" });
      router.push(`/files/${encodeURIComponent(created.id)}`);
    } catch (error) {
      notify({
        title: "Couldn't create the container",
        body: error instanceof Error ? error.message : "Unknown error",
        intent: "error",
      });
      throw error;
    }
  };

  return (
    <AuthBoundary>
      <AppShell>
        <section className={styles.stack}>
          <div className={styles.header}>
            <div>
              <Text as="h1" size={800} weight="semibold">Files</Text>
              <Text block>Choose a SharePoint Embedded container.</Text>
            </div>
            <Button appearance="primary" icon={<AddRegular />} disabled={create.isPending} onClick={() => setDialogOpen(true)}>
              {create.isPending ? "Creating…" : "Create container"}
            </Button>
          </div>
          {containers.isLoading ? <Spinner label="Loading containers" /> : null}
          {containers.data && containers.data.length > 1 ? (
            <div className={styles.cards}>
              <ContainerCard container={containers.data[0]} featured />
              <div className={styles.sideCards}>{containers.data.slice(1).map((container) => <ContainerCard key={container.id} container={container} />)}</div>
            </div>
          ) : null}
          {containers.data?.length === 0 ? (
            <Card className={styles.empty} appearance="outline">
              <FolderRegular fontSize={32} aria-label="No containers" />
              <div>
                <Text weight="semibold" size={500}>No containers yet.</Text>
                <Text block>Create your first workspace to start uploading files.</Text>
              </div>
              <Button appearance="primary" icon={<AddRegular />} disabled={create.isPending} onClick={() => setDialogOpen(true)}>
                {create.isPending ? "Creating…" : "Create container"}
              </Button>
            </Card>
          ) : null}
        </section>
        <NewContainerDialog open={dialogOpen} onOpenChange={setDialogOpen} onCreate={handleCreate} />
      </AppShell>
    </AuthBoundary>
  );
}
