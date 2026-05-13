import {
  DocumentPdfRegular,
  DocumentRegular,
  DocumentTableRegular,
  DocumentTextRegular,
  DocumentWordRegular,
  FolderRegular,
  ImageRegular,
  SlideTextRegular,
} from "@fluentui/react-icons";
import type { DriveItem } from "@/lib/api/types";
import { extensionFor } from "@/lib/files";

interface FileIconProps {
  item?: Pick<DriveItem, "name" | "folder" | "file">;
  mimeType?: string;
  name?: string;
}

export function FileIcon({ item, mimeType, name }: FileIconProps) {
  const fileName = item?.name ?? name ?? "";
  const mime = (item?.file?.mimeType ?? mimeType ?? "").toLowerCase();
  const ext = extensionFor(fileName);
  const label = item?.folder ? "Folder" : `${ext || "File"} file`;

  if (item?.folder) return <FolderRegular aria-label={label} />;
  if (["doc", "docx", "dotx"].includes(ext) || mime.includes("word")) return <DocumentWordRegular aria-label="Word document" />;
  if (["xls", "xlsx", "xlsm", "csv"].includes(ext) || mime.includes("spreadsheet") || mime.includes("excel")) return <DocumentTableRegular aria-label="Excel workbook" />;
  if (["ppt", "pptx"].includes(ext) || mime.includes("presentation") || mime.includes("powerpoint")) return <SlideTextRegular aria-label="PowerPoint presentation" />;
  if (ext === "pdf" || mime.includes("pdf")) return <DocumentPdfRegular aria-label="PDF document" />;
  if (["png", "jpg", "jpeg", "gif", "webp", "svg"].includes(ext) || mime.startsWith("image/")) return <ImageRegular aria-label="Image file" />;
  if (["txt", "md", "rtf"].includes(ext) || mime.startsWith("text/")) return <DocumentTextRegular aria-label="Text file" />;
  return <DocumentRegular aria-label={label} />;
}
