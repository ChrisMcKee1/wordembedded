export interface Container {
  id: string;
  displayName: string;
  driveId: string;
  createdDateTime: string;
  status: string;
}

export interface IdentityUser {
  displayName?: string;
  email?: string;
}

export interface ModifiedBy {
  user?: IdentityUser;
  displayName?: string;
  email?: string;
}

export interface DriveItem {
  id: string;
  name: string;
  webUrl: string;
  size: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
  lastModifiedBy?: ModifiedBy;
  folder?: { childCount: number };
  file?: { mimeType: string };
  parentReference?: { id: string; path: string; driveId: string };
}

export interface DriveChildrenResponse {
  value: DriveItem[];
}

export interface PreviewResponse {
  getUrl: string;
}

export type LinkType = "view" | "edit" | "embed";
export type LinkScope = "organization" | "users";
export type ShareRole = "read" | "write";

export interface CreateLinkRequest {
  type: LinkType;
  scope: LinkScope;
  password?: string;
  expirationDateTime?: string;
  recipients?: { email: string }[];
}

export interface CreateLinkResponse {
  link: { webUrl: string };
  id?: string;
  roles?: string[];
}

export interface InviteRequest {
  recipients: { email: string }[];
  roles: ShareRole[];
  sendInvitation: boolean;
  message?: string;
}

export interface Permission {
  id: string;
  roles?: string[];
  link?: { type?: string; scope?: string; webUrl?: string };
  grantedTo?: { user?: IdentityUser };
  grantedToV2?: { user?: IdentityUser };
  invitation?: { email?: string; signInRequired?: boolean };
}

export interface Version {
  id: string;
  lastModifiedDateTime: string;
  size: number;
  lastModifiedBy?: ModifiedBy;
  downloadUrl?: string;
}

export interface VersionsResponse {
  value: Version[];
}

export interface UploadSessionResponse {
  uploadUrl: string;
  expirationDateTime: string;
}

export interface SearchHit {
  id: string;
  name: string;
  webUrl?: string;
  driveId: string;
  containerId?: string;
  containerName?: string;
  size?: number;
  lastModifiedDateTime?: string;
  parentReference?: { id?: string; path?: string; driveId?: string };
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  snippet?: string;
}

export interface SearchResponse {
  value: SearchHit[];
}

export interface RecycleBinItem {
  id: string;
  name?: string;
  size?: number;
  deletedDateTime?: string;
  deletedBy?: { displayName?: string; email?: string };
  lastModifiedDateTime?: string;
  originalDriveItemId?: string;
}

export interface RecycleBinResponse {
  value: RecycleBinItem[];
}

export interface ThumbnailImage {
  url?: string;
  width?: number;
  height?: number;
}

export interface Thumbnail {
  id?: string;
  small?: ThumbnailImage;
  medium?: ThumbnailImage;
  large?: ThumbnailImage;
}

export interface ThumbnailSetResponse {
  value: Thumbnail[];
}

