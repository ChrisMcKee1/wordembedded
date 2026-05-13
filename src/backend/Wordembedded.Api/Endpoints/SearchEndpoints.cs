using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Search.Query;
using Wordembedded.Api.Contracts;
using Wordembedded.Api.Services;

namespace Wordembedded.Api.Endpoints;

public static class SearchEndpoints
{
    public static IEndpointRouteBuilder MapSearchEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api/search").RequireAuthorization().WithTags("Search");

        group.MapGet(string.Empty, async (
            [FromQuery(Name = "q")] string query,
            [FromQuery] string? driveId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return EndpointResults.BadRequestProblem("Query is required", "Provide a non-empty q query string parameter.");
            }

            var response = await graphClient.Search.Query.PostAsQueryPostResponseAsync(new QueryPostRequestBody
            {
                Requests = new List<SearchRequest>
                {
                    new()
                    {
                        EntityTypes = new List<EntityType?> { EntityType.DriveItem },
                        Query = new SearchQuery { QueryString = $"{query} AND ContentTypeId:0x0101*" },
                        From = 0,
                        Size = 25
                    }
                }
            }, cancellationToken: cancellationToken);

            var hits = response?.Value?
                .SelectMany(searchResponse => searchResponse.HitsContainers ?? [])
                .SelectMany(container => container.Hits ?? [])
                .Select(hit =>
                {
                    var item = hit.Resource as DriveItem;
                    return new SearchHitResponse(
                        hit.HitId,
                        hit.Rank,
                        hit.Summary,
                        item is null ? null : GraphDtoMapper.ToDriveItem(item));
                })
                .Where(hit => string.IsNullOrWhiteSpace(driveId) || string.Equals(hit.Resource?.ParentReference?.DriveId, driveId, StringComparison.OrdinalIgnoreCase))
                .ToList() ?? [];

            return Results.Ok(new { value = hits });
        });

        return app;
    }
}
