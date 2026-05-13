using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Wordembedded.Api.Contracts;

namespace Wordembedded.Api.Endpoints;

public static class ThumbnailEndpoints
{
    public static IEndpointRouteBuilder MapThumbnailEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api/drives/{driveId}/items/{itemId}/thumbnails").RequireAuthorization().WithTags("Thumbnails");

        group.MapGet(string.Empty, async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            var set = await graphClient.Drives[driveId].Items[itemId].Thumbnails.GetAsync(cancellationToken: cancellationToken);
            var thumbs = (set?.Value ?? []).Select(thumb => new ThumbnailResponse(
                thumb.Id,
                thumb.Small is null ? null : new ThumbnailImage(thumb.Small.Url, thumb.Small.Width, thumb.Small.Height),
                thumb.Medium is null ? null : new ThumbnailImage(thumb.Medium.Url, thumb.Medium.Width, thumb.Medium.Height),
                thumb.Large is null ? null : new ThumbnailImage(thumb.Large.Url, thumb.Large.Width, thumb.Large.Height)));

            return Results.Ok(new { value = thumbs });
        });

        return app;
    }
}
