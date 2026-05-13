import { Card, ProgressBar, Text, makeStyles, tokens } from "@fluentui/react-components";

export interface UploadProgressItem {
  id: string;
  name: string;
  progress: number;
  status: "uploading" | "done" | "error";
}

const useStyles = makeStyles({
  list: {
    display: "grid",
    gap: tokens.spacingVerticalS,
  },
  row: {
    padding: tokens.spacingVerticalS,
    display: "grid",
    gap: tokens.spacingVerticalXS,
  },
});

export function UploadProgress({ uploads }: { uploads: UploadProgressItem[] }) {
  const styles = useStyles();
  if (uploads.length === 0) return null;

  return (
    <div className={styles.list} aria-live="polite">
      {uploads.map((upload) => (
        <Card key={upload.id} className={styles.row} appearance="outline">
          <Text weight="semibold">{upload.name}</Text>
          <ProgressBar value={upload.progress} max={100} thickness="medium" />
          <Text size={200}>{upload.status === "error" ? "Upload failed" : upload.status === "done" ? "Uploaded" : `${Math.round(upload.progress)}%`}</Text>
        </Card>
      ))}
    </div>
  );
}
