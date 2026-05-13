namespace Wordembedded.Api.Contracts;

public sealed record ContainerSummaryResponse(
    string? Id,
    string? DisplayName,
    string? DriveId,
    DateTimeOffset? CreatedDateTime,
    string? Status);

public sealed record ContainerDetailResponse(
    string? Id,
    string? DisplayName,
    string? Description,
    Guid? ContainerTypeId,
    string? DriveId,
    DateTimeOffset? CreatedDateTime,
    string? Status);

public sealed record CreateContainerRequest(string DisplayName, string? Description);

public sealed record DriveItemResponse(
    string? Id,
    string? Name,
    string? WebUrl,
    long? Size,
    DateTimeOffset? LastModifiedDateTime,
    IdentityResponse? LastModifiedBy,
    DateTimeOffset? CreatedDateTime,
    FolderResponse? Folder,
    FileResponse? File,
    ParentReferenceResponse? ParentReference);

public sealed record IdentityResponse(string? DisplayName, string? Email);

public sealed record FolderResponse(int? ChildCount);

public sealed record FileResponse(string? MimeType);

public sealed record ParentReferenceResponse(string? Id, string? Path, string? DriveId);

public sealed record UpdateDriveItemRequest(string? Name, ParentReferenceRequest? ParentReference);

public sealed record ParentReferenceRequest(string? Id, string? Path, string? DriveId);

public sealed record CreateFolderRequest(string Name);

public sealed record CreateUploadSessionRequest(string Name, long? Size, string? ConflictBehavior);

public sealed record UploadSessionResponse(string? UploadUrl, DateTimeOffset? ExpirationDateTime);

public sealed record PreviewRequest(string? Page, double? Zoom, string? Viewer, bool? AllowEdit, bool? Chromeless);

public sealed record PreviewResponse(string? GetUrl, string? PostUrl, string? PostParameters);

public sealed record RecipientRequest(string Email);

public sealed record CreateLinkRequest(
    string Type,
    string Scope,
    string? Password,
    DateTimeOffset? ExpirationDateTime,
    IReadOnlyList<RecipientRequest>? Recipients);

public sealed record SharingLinkResponse(
    string? Id,
    IReadOnlyList<string>? Roles,
    string? LinkType,
    string? Scope,
    string? WebUrl,
    DateTimeOffset? ExpirationDateTime);

public sealed record PermissionResponse(
    string? Id,
    IReadOnlyList<string>? Roles,
    string? LinkType,
    string? Scope,
    string? WebUrl,
    DateTimeOffset? ExpirationDateTime,
    IReadOnlyList<IdentityResponse>? GrantedTo);

public sealed record InviteRequest(
    IReadOnlyList<RecipientRequest> Recipients,
    IReadOnlyList<string> Roles,
    bool SendInvitation,
    string? Message);

public sealed record VersionResponse(
    string? Id,
    DateTimeOffset? LastModifiedDateTime,
    IdentityResponse? LastModifiedBy,
    long? Size);

public sealed record SearchHitResponse(
    string? HitId,
    int? Rank,
    string? Summary,
    DriveItemResponse? Resource);

public sealed record CreateSubscriptionRequest(string DriveId);

public sealed record GraphSubscriptionRecord(
    string Id,
    string DriveId,
    string Resource,
    DateTimeOffset ExpirationDateTime,
    string ClientState);

public sealed record GraphWebhookPayload(IReadOnlyList<GraphWebhookNotification> Value);

public sealed record GraphWebhookNotification(
    string? SubscriptionId,
    string? ClientState,
    string? ChangeType,
    string? Resource,
    GraphWebhookResourceData? ResourceData,
    string? TenantId);

public sealed record GraphWebhookResourceData(string? Id);

public sealed record RecycleBinItemResponse(
    string? Id,
    string? Name,
    long? Size,
    DateTimeOffset? DeletedDateTime,
    IdentityResponse? DeletedBy,
    DateTimeOffset? LastModifiedDateTime,
    string? OriginalDriveItemId);

public sealed record RecycleBinActionRequest(IReadOnlyList<string> Ids);

public sealed record ThumbnailResponse(
    string? Id,
    ThumbnailImage? Small,
    ThumbnailImage? Medium,
    ThumbnailImage? Large);

public sealed record ThumbnailImage(string? Url, int? Width, int? Height);

public sealed record SharePointSearchHit(
    string? HitId,
    string? Id,
    string? Name,
    string? WebUrl,
    long? Size,
    string? LastModifiedDateTime,
    string? DriveId,
    string? ParentPath,
    string? SiteId,
    string? Summary);

public sealed record SharePointSiteSummary(string? Id, string? DisplayName, string? Name, string? WebUrl);

public sealed record SharePointDriveSummary(string? Id, string? Name, string? WebUrl, string? DriveType);

public sealed record SharePointBrowseItem(string? Id, string? Name, string? WebUrl, long? Size, string? LastModifiedDateTime, bool IsFolder, string? MimeType);

public sealed record ImportFileRequest(string SourceDriveId, string SourceItemId, string TargetDriveId, string? TargetFolderId, string? NewName);

public sealed record ImportResponse(string? MonitorUrl, string? Status);

public sealed record PublishRequest(string TargetDriveId, string? TargetFolderId, string? NewName, string? ConflictBehavior);

public sealed record PublishedReference(
    string? DriveId,
    string? ItemId,
    string? WebUrl,
    string? Name,
    DateTimeOffset? PublishedAt);

public sealed record PublishResponse(string? MonitorUrl, string? Status, PublishedReference? Published);

public sealed record WorkflowStatusResponse(string Status, PublishedReference? Published);

public sealed record SetWorkflowRequest(string Status);

public sealed record PublishTargetResponse(string? SiteId, string? DriveId, string? FolderId, string? FolderPath);
