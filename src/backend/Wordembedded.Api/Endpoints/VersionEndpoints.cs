using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Wordembedded.Api.Services;

namespace Wordembedded.Api.Endpoints;

public static class VersionEndpoints
{
    public static IEndpointRouteBuilder MapVersionEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api/drives/{driveId}/items/{itemId}/versions").RequireAuthorization().WithTags("Versions");

        group.MapGet(string.Empty, async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            var versions = await graphClient.Drives[driveId].Items[itemId].Versions.GetAsync(cancellationToken: cancellationToken);
            return Results.Ok(new { value = (versions?.Value ?? []).Select(GraphDtoMapper.ToVersion) });
        });

        group.MapPost("/{versionId}/restore", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            [FromRoute] string versionId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            await graphClient.Drives[driveId]
                .Items[itemId]
                .Versions[versionId]
                .RestoreVersion
                .PostAsync(cancellationToken: cancellationToken);

            return Results.NoContent();
        });

        return app;
    }
}
