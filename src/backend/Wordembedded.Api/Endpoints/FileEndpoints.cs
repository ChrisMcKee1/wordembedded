using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Wordembedded.Api.Contracts;
using Wordembedded.Api.Services;
using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;

namespace Wordembedded.Api.Endpoints;

public static class FileEndpoints
{
    private const long MaxSmallUploadBytes = 4 * 1024 * 1024;

    public static IEndpointRouteBuilder MapFileEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api/drives/{driveId}").RequireAuthorization().WithTags("Files");

        group.MapGet("/root/children", async (
            [FromRoute] string driveId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            var children = await graphClient.Drives[driveId].Items["root"].Children.GetAsync(cancellationToken: cancellationToken);
            return Results.Ok(new { value = (children?.Value ?? []).Select(GraphDtoMapper.ToDriveItem) });
        });

        group.MapGet("/items/{itemId}/children", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            var children = await graphClient.Drives[driveId].Items[itemId].Children.GetAsync(cancellationToken: cancellationToken);
            return Results.Ok(new { value = (children?.Value ?? []).Select(GraphDtoMapper.ToDriveItem) });
        });

        group.MapGet("/items/{itemId}", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            var item = await graphClient.Drives[driveId].Items[itemId].GetAsync(cancellationToken: cancellationToken);
            return item is null
                ? EndpointResults.NotFoundProblem("Drive item not found", $"Drive item '{itemId}' was not found.")
                : Results.Ok(GraphDtoMapper.ToDriveItem(item));
        });

        group.MapPatch("/items/{itemId}", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            [FromBody] UpdateDriveItemRequest request,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            if (string.IsNullOrWhiteSpace(request.Name) && request.ParentReference is null)
            {
                return EndpointResults.BadRequestProblem("No update specified", "Provide name and/or parentReference.");
            }

            var updated = await graphClient.Drives[driveId].Items[itemId].PatchAsync(new DriveItem
            {
                Name = string.IsNullOrWhiteSpace(request.Name) ? null : request.Name,
                ParentReference = request.ParentReference is null ? null : GraphDtoMapper.ToItemReference(request.ParentReference)
            }, cancellationToken: cancellationToken);

            return updated is null
                ? EndpointResults.NotFoundProblem("Drive item not found", $"Drive item '{itemId}' was not found.")
                : Results.Ok(GraphDtoMapper.ToDriveItem(updated));
        });

        group.MapDelete("/items/{itemId}", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            await graphClient.Drives[driveId].Items[itemId].DeleteAsync(cancellationToken: cancellationToken);
            return Results.NoContent();
        });

        group.MapPost("/items/{parentId}/folders", async (
            [FromRoute] string driveId,
            [FromRoute] string parentId,
            [FromBody] CreateFolderRequest request,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            if (string.IsNullOrWhiteSpace(request.Name))
            {
                return EndpointResults.BadRequestProblem("Folder name is required", "Folder name cannot be empty.");
            }

            var folder = await graphClient.Drives[driveId].Items[parentId].Children.PostAsync(new DriveItem
            {
                Name = request.Name,
                Folder = new Folder(),
                AdditionalData = new Dictionary<string, object>
                {
                    ["@microsoft.graph.conflictBehavior"] = "rename"
                }
            }, cancellationToken: cancellationToken);

            return folder is null
                ? EndpointResults.NotFoundProblem("Folder creation failed", "Microsoft Graph did not return the created folder.")
                : Results.Ok(GraphDtoMapper.ToDriveItem(folder));
        });

        group.MapPost("/items/{parentId}/upload", async (
            [FromRoute] string driveId,
            [FromRoute] string parentId,
            HttpRequest httpRequest,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            if (!httpRequest.HasFormContentType)
            {
                return EndpointResults.BadRequestProblem("Invalid content type", "Use multipart/form-data with a file field.");
            }

            var form = await httpRequest.ReadFormAsync(cancellationToken);
            var file = form.Files.GetFile("file") ?? form.Files.FirstOrDefault();
            if (file is null || file.Length == 0)
            {
                return EndpointResults.BadRequestProblem("File is required", "Upload a non-empty file in the multipart form.");
            }

            if (file.Length > MaxSmallUploadBytes)
            {
                return EndpointResults.BadRequestProblem("File too large", "Small uploads are limited to 4 MB. Use upload-session for larger files.");
            }

            var fileName = Path.GetFileName(file.FileName);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return EndpointResults.BadRequestProblem("File name is required", "The uploaded file must include a file name.");
            }

            await using var stream = file.OpenReadStream();
            var uploaded = await graphClient.Drives[driveId]
                .Items[parentId]
                .ItemWithPath(fileName)
                .Content
                .PutAsync(stream, cancellationToken: cancellationToken);

            return uploaded is null
                ? EndpointResults.NotFoundProblem("Upload failed", "Microsoft Graph did not return the uploaded item.")
                : Results.Ok(GraphDtoMapper.ToDriveItem(uploaded));
        }).DisableAntiforgery();

        group.MapPost("/items/{parentId}/upload-session", async (
            [FromRoute] string driveId,
            [FromRoute] string parentId,
            [FromBody] CreateUploadSessionRequest request,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            if (string.IsNullOrWhiteSpace(request.Name))
            {
                return EndpointResults.BadRequestProblem("File name is required", "Upload session name cannot be empty.");
            }

            var conflictBehavior = string.IsNullOrWhiteSpace(request.ConflictBehavior) ? "rename" : request.ConflictBehavior;
            var itemProperties = new DriveItemUploadableProperties
            {
                Name = request.Name,
                AdditionalData = new Dictionary<string, object>
                {
                    ["@microsoft.graph.conflictBehavior"] = conflictBehavior
                }
            };

            if (request.Size is not null)
            {
                itemProperties.AdditionalData["fileSize"] = request.Size.Value;
            }

            var uploadSession = await graphClient.Drives[driveId]
                .Items[parentId]
                .ItemWithPath(request.Name)
                .CreateUploadSession
                .PostAsync(new DriveUpload.CreateUploadSessionPostRequestBody
                {
                    Item = itemProperties
                }, cancellationToken: cancellationToken);

            return uploadSession is null
                ? EndpointResults.NotFoundProblem("Upload session creation failed", "Microsoft Graph did not return an upload session.")
                : Results.Ok(new UploadSessionResponse(uploadSession.UploadUrl, uploadSession.ExpirationDateTime));
        });

        return app;
    }
}
