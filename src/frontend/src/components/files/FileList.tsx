"use client";

import {
  Avatar,
  Button,
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Text,
  Tooltip,
  createTableColumn,
  makeStyles,
  tokens,
  type TableColumnDefinition,
  type TableRowId,
} from "@fluentui/react-components";
import { DeleteRegular, EditRegular, HistoryRegular, MoreHorizontalRegular, OpenRegular, RenameRegular, ShareRegular } from "@fluentui/react-icons";
import { useRouter } from "next/navigation";
import { FileIcon } from "./FileIcon";
import type { DriveItem } from "@/lib/api/types";
import type { RoutePathSegment } from "@/lib/files";
import { encodeRouteSegment, formatBytes, formatDate, modifiedByName, supportsOfficeEdit } from "@/lib/files";

const useStyles = makeStyles({
  nameCell: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    minWidth: 0,
  },
  nameText: {
    overflowX: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  person: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  rowActions: {
    display: "flex",
    justifyContent: "flex-end",
  },
  empty: {
    paddingTop: tokens.spacingVerticalXXL,
    paddingBottom: tokens.spacingVerticalXXL,
    textAlign: "center",
    color: tokens.colorNeutralForeground2,
  },
});

interface FileListProps {
  containerId: string;
  driveId: string;
  items: DriveItem[];
  pathSegments: RoutePathSegment[];
  selectedItems: Set<TableRowId>;
  onSelectionChange: (ids: Set<TableRowId>) => void;
  onRename: (item: DriveItem) => void;
  onDelete: (item: DriveItem) => void;
  onShare: (item: DriveItem) => void;
  onVersions: (item: DriveItem) => void;
}

export function FileList({ containerId, driveId, items, pathSegments, selectedItems, onSelectionChange, onRename, onDelete, onShare, onVersions }: FileListProps) {
  const styles = useStyles();
  const router = useRouter();

  const openItem = (item: DriveItem) => {
    if (item.folder) {
      const nextPath = [...pathSegments, { id: item.id, name: item.name }].map(encodeRouteSegment).join("/");
      router.push(`/files/${encodeURIComponent(containerId)}/${nextPath}`);
      return;
    }
    router.push(`/files/${encodeURIComponent(containerId)}/preview/${encodeURIComponent(item.id)}?driveId=${encodeURIComponent(driveId)}`);
  };

  const columns: TableColumnDefinition<DriveItem>[] = [
    createTableColumn<DriveItem>({
      columnId: "name",
      renderHeaderCell: () => "Name",
      renderCell: (item) => (
        <Button appearance="transparent" onClick={() => openItem(item)} className={styles.nameCell}>
          <FileIcon item={item} />
          <span className={styles.nameText}>{item.name}</span>
        </Button>
      ),
    }),
    createTableColumn<DriveItem>({
      columnId: "modified",
      renderHeaderCell: () => "Modified",
      renderCell: (item) => <Text>{formatDate(item.lastModifiedDateTime)}</Text>,
    }),
    createTableColumn<DriveItem>({
      columnId: "modifiedBy",
      renderHeaderCell: () => "Modified by",
      renderCell: (item) => {
        const name = modifiedByName(item.lastModifiedBy);
        return (
          <div className={styles.person}>
            <Avatar name={name} size={24} />
            <Text>{name}</Text>
          </div>
        );
      },
    }),
    createTableColumn<DriveItem>({
      columnId: "size",
      renderHeaderCell: () => "Size",
      renderCell: (item) => <Text>{item.folder ? `${item.folder.childCount} items` : formatBytes(item.size)}</Text>,
    }),
    createTableColumn<DriveItem>({
      columnId: "actions",
      renderHeaderCell: () => "Actions",
      renderCell: (item) => (
        <div className={styles.rowActions}>
          <Menu positioning="below-end">
            <MenuTrigger disableButtonEnhancement>
              <Tooltip content={`Actions for ${item.name}`} relationship="label">
                <Button appearance="subtle" icon={<MoreHorizontalRegular />} aria-label={`Actions for ${item.name}`} />
              </Tooltip>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem icon={<OpenRegular />} onClick={() => openItem(item)}>{item.folder ? "Open" : "Preview"}</MenuItem>
                {!item.folder && supportsOfficeEdit(item) ? <MenuItem icon={<EditRegular />} onClick={() => window.open(item.webUrl, "_blank", "noopener,noreferrer")}>Edit in Office</MenuItem> : null}
                {!item.folder ? <MenuItem icon={<ShareRegular />} onClick={() => onShare(item)}>Share</MenuItem> : null}
                <MenuItem icon={<RenameRegular />} onClick={() => onRename(item)}>Rename</MenuItem>
                {!item.folder ? <MenuItem icon={<HistoryRegular />} onClick={() => onVersions(item)}>Versions</MenuItem> : null}
                <MenuItem icon={<DeleteRegular />} onClick={() => onDelete(item)}>Delete</MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        </div>
      ),
    }),
  ];

  if (items.length === 0) {
    return <div className={styles.empty}>No files yet. Drop files here or click Upload.</div>;
  }

  return (
    <DataGrid
      items={items}
      columns={columns}
      getRowId={(item) => item.id}
      selectionMode="multiselect"
      selectedItems={selectedItems}
      onSelectionChange={(_, data) => onSelectionChange(new Set(data.selectedItems))}
      sortable
      resizableColumns
    >
      <DataGridHeader>
        <DataGridRow selectionCell={{ checkboxIndicator: { "aria-label": "Select all files" } }}>
          {({ renderHeaderCell }) => <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>}
        </DataGridRow>
      </DataGridHeader>
      <DataGridBody<DriveItem>>
        {({ item, rowId }) => (
          <DataGridRow<DriveItem> key={rowId} selectionCell={{ checkboxIndicator: { "aria-label": `Select ${item.name}` } }}>
            {({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
          </DataGridRow>
        )}
      </DataGridBody>
    </DataGrid>
  );
}
