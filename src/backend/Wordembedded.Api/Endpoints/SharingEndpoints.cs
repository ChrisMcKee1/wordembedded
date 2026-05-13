using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Wordembedded.Api.Contracts;
using Wordembedded.Api.Services;
using CreateLink = Microsoft.Graph.Drives.Item.Items.Item.CreateLink;
using Invite = Microsoft.Graph.Drives.Item.Items.Item.Invite;

namespace Wordembedded.Api.Endpoints;

public static class SharingEndpoints
{
    private static readonly HashSet<string> AllowedLinkTypes = new(StringComparer.OrdinalIgnoreCase) { "view", "edit", "embed" };
    private static readonly HashSet<string> AllowedLinkScopes = new(StringComparer.OrdinalIgnoreCase) { "organization", "users" };
    private static readonly HashSet<string> AllowedInviteRoles = new(StringComparer.OrdinalIgnoreCase) { "read", "write" };

    public static IEndpointRouteBuilder MapSharingEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api/drives/{driveId}/items/{itemId}").RequireAuthorization().WithTags("Sharing");

        group.MapPost("/createLink", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            [FromBody] CreateLinkRequest request,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            if (string.Equals(request.Scope, "anonymous", StringComparison.OrdinalIgnoreCase))
            {
                return EndpointResults.BadRequestProblem("Anonymous links are disabled", "Phase 1 does not allow anonymous sharing links.");
            }

            if (!AllowedLinkTypes.Contains(request.Type) || !AllowedLinkScopes.Contains(request.Scope))
            {
                return EndpointResults.BadRequestProblem("Invalid sharing link", "type must be view, edit, or embed; scope must be organization or users.");
            }

            var body = new CreateLink.CreateLinkPostRequestBody
            {
                Type = request.Type,
                Scope = request.Scope,
                Password = request.Password,
                ExpirationDateTime = request.ExpirationDateTime
            };

            if (request.Recipients is { Count: > 0 })
            {
                body.AdditionalData = new Dictionary<string, object>
                {
                    ["recipients"] = request.Recipients.Select(GraphDtoMapper.ToDriveRecipient).ToList()
                };
            }

            var permission = await graphClient.Drives[driveId]
                .Items[itemId]
                .CreateLink
                .PostAsync(body, cancellationToken: cancellationToken);

            return permission is null
                ? EndpointResults.NotFoundProblem("Link creation failed", "Microsoft Graph did not return a permission.")
                : Results.Ok(GraphDtoMapper.ToSharingLink(permission));
        });

        group.MapGet("/permissions", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            var permissions = await graphClient.Drives[driveId].Items[itemId].Permissions.GetAsync(cancellationToken: cancellationToken);
            return Results.Ok((permissions?.Value ?? []).Select(GraphDtoMapper.ToPermission));
        });

        group.MapPost("/invite", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            [FromBody] InviteRequest request,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            if (request.Recipients.Count == 0 || request.Recipients.Any(r => string.IsNullOrWhiteSpace(r.Email)))
            {
                return EndpointResults.BadRequestProblem("Recipients are required", "Provide at least one recipient email address.");
            }

            if (request.Roles.Count == 0 || request.Roles.Any(role => !AllowedInviteRoles.Contains(role)))
            {
                return EndpointResults.BadRequestProblem("Invalid roles", "roles must include read and/or write.");
            }

            var permissions = await graphClient.Drives[driveId]
                .Items[itemId]
                .Invite
                .PostAsInvitePostResponseAsync(new Invite.InvitePostRequestBody
                {
                    Recipients = request.Recipients.Select(GraphDtoMapper.ToDriveRecipient).ToList(),
                    Roles = request.Roles.ToList(),
                    SendInvitation = request.SendInvitation,
                    Message = request.Message
                }, cancellationToken: cancellationToken);

            return Results.Ok((permissions?.Value ?? []).Select(GraphDtoMapper.ToPermission));
        });

        group.MapDelete("/permissions/{permissionId}", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            [FromRoute] string permissionId,
            GraphServiceClient graphClient,
            CancellationToken cancellationToken) =>
        {
            await graphClient.Drives[driveId].Items[itemId].Permissions[permissionId].DeleteAsync(cancellationToken: cancellationToken);
            return Results.NoContent();
        });

        return app;
    }
}
