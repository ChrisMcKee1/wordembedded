using Microsoft.Graph.Models;
using Wordembedded.Api.Contracts;

namespace Wordembedded.Api.Services;

public static class GraphDtoMapper
{
    public static ContainerSummaryResponse ToContainerSummary(FileStorageContainer container, string? driveId) => new(
        container.Id,
        container.DisplayName,
        driveId,
        container.CreatedDateTime,
        container.Status?.ToString());

    public static ContainerDetailResponse ToContainerDetail(FileStorageContainer container, string? driveId) => new(
        container.Id,
        container.DisplayName,
        container.Description,
        container.ContainerTypeId,
        driveId,
        container.CreatedDateTime,
        container.Status?.ToString());

    public static DriveItemResponse ToDriveItem(DriveItem item) => new(
        item.Id,
        item.Name,
        item.WebUrl,
        item.Size,
        item.LastModifiedDateTime,
        ToIdentity(item.LastModifiedBy),
        item.CreatedDateTime,
        item.Folder is null ? null : new FolderResponse(item.Folder.ChildCount),
        item.File is null ? null : new FileResponse(item.File.MimeType),
        item.ParentReference is null ? null : new ParentReferenceResponse(
            item.ParentReference.Id,
            item.ParentReference.Path,
            item.ParentReference.DriveId));

    public static VersionResponse ToVersion(DriveItemVersion version) => new(
        version.Id,
        version.LastModifiedDateTime,
        ToIdentity(version.LastModifiedBy),
        version.Size);

    public static PermissionResponse ToPermission(Permission permission)
    {
        var identities = new List<IdentityResponse>();

        if (permission.GrantedToV2 is not null)
        {
            var identity = ToIdentity(permission.GrantedToV2);
            if (identity is not null)
            {
                identities.Add(identity);
            }
        }

        if (permission.GrantedToIdentitiesV2 is not null)
        {
            identities.AddRange(permission.GrantedToIdentitiesV2.Select(ToIdentity).OfType<IdentityResponse>());
        }

        return new PermissionResponse(
            permission.Id,
            permission.Roles,
            permission.Link?.Type?.ToString(),
            permission.Link?.Scope?.ToString(),
            permission.Link?.WebUrl,
            permission.ExpirationDateTime,
            identities.Count == 0 ? null : identities);
    }

    public static SharingLinkResponse ToSharingLink(Permission permission) => new(
        permission.Id,
        permission.Roles,
        permission.Link?.Type?.ToString(),
        permission.Link?.Scope?.ToString(),
        permission.Link?.WebUrl,
        permission.ExpirationDateTime);

    public static ItemReference ToItemReference(ParentReferenceRequest request) => new()
    {
        Id = request.Id,
        Path = request.Path,
        DriveId = request.DriveId
    };

    public static DriveRecipient ToDriveRecipient(RecipientRequest recipient) => new()
    {
        Email = recipient.Email
    };

    private static IdentityResponse? ToIdentity(IdentitySet? identitySet)
    {
        var identity = identitySet?.User ?? identitySet?.Application ?? identitySet?.Device;
        if (identity is null)
        {
            return null;
        }

        return new IdentityResponse(identity.DisplayName, TryGetString(identity.AdditionalData, "email"));
    }

    private static string? TryGetString(IDictionary<string, object>? additionalData, string key)
    {
        if (additionalData is null || !additionalData.TryGetValue(key, out var value))
        {
            return null;
        }

        return value?.ToString();
    }
}
