using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Wordembedded.Api.Contracts;
using Wordembedded.Api.Options;
using Wordembedded.Api.Services;

namespace Wordembedded.Api.Endpoints;

public static class ContainerEndpoints
{
    public static IEndpointRouteBuilder MapContainerEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api/containers").RequireAuthorization().WithTags("Containers");

        group.MapGet(string.Empty, async (
            GraphServiceClient graphClient,
            IOptions<SharePointEmbeddedOptions> options,
            CancellationToken cancellationToken) =>
        {
            var containerTypeId = options.Value.ContainerTypeId;
            var containers = await graphClient.Storage.FileStorage.Containers.GetAsync(request =>
            {
                request.QueryParameters.Filter = $"containerTypeId eq {containerTypeId}";
            }, cancellationToken);

            var responses = new List<ContainerSummaryResponse>();
            foreach (var container in containers?.Value ?? [])
            {
                var driveId = await GetDriveIdAsync(graphClient, container.Id, cancellationToken);
                responses.Add(GraphDtoMapper.ToContainerSummary(container, driveId));
            }

            return Results.Ok(responses);
        });

        group.MapGet("/{containerId}", async (
            [FromRoute] string containerId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            var container = await graphClient.Storage.FileStorage.Containers[containerId].GetAsync(cancellationToken: cancellationToken);
            if (container is null)
            {
                return EndpointResults.NotFoundProblem("Container not found", $"Container '{containerId}' was not found.");
            }

            var driveId = await GetDriveIdAsync(graphClient, container.Id, cancellationToken);
            return Results.Ok(GraphDtoMapper.ToContainerDetail(container, driveId));
        });

        group.MapPost(string.Empty, async (
            [FromBody] CreateContainerRequest request,
            GraphServiceClient graphClient,
            IOptions<SharePointEmbeddedOptions> options,
            CancellationToken cancellationToken) =>
        {
            if (string.IsNullOrWhiteSpace(request.DisplayName))
            {
                return EndpointResults.BadRequestProblem("Display name is required", "Container displayName cannot be empty.");
            }

            if (!Guid.TryParse(options.Value.ContainerTypeId, out var containerTypeId))
            {
                return EndpointResults.BadRequestProblem("Invalid container type", "SharePointEmbedded:ContainerTypeId must be a valid GUID.");
            }

            var container = await graphClient.Storage.FileStorage.Containers.PostAsync(new FileStorageContainer
            {
                DisplayName = request.DisplayName,
                Description = request.Description,
                ContainerTypeId = containerTypeId
            }, cancellationToken: cancellationToken);

            if (container is null)
            {
                return EndpointResults.NotFoundProblem("Container creation failed", "Microsoft Graph did not return the created container.");
            }

            var driveId = await GetDriveIdAsync(graphClient, container.Id, cancellationToken);
            return Results.Ok(GraphDtoMapper.ToContainerDetail(container, driveId));
        });

        return app;
    }

    private static async Task<string?> GetDriveIdAsync(GraphServiceClient graphClient, string? containerId, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(containerId))
        {
            return null;
        }

        var drive = await graphClient.Storage.FileStorage.Containers[containerId].Drive.GetAsync(cancellationToken: cancellationToken);
        return drive?.Id;
    }
}
