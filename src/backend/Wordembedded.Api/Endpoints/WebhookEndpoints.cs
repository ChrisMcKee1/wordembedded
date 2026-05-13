using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Microsoft.Graph.Models;
using Wordembedded.Api.Contracts;
using Wordembedded.Api.Options;
using Wordembedded.Api.Services;

namespace Wordembedded.Api.Endpoints;

public static class WebhookEndpoints
{
    private static readonly TimeSpan SubscriptionLifetime = TimeSpan.FromDays(2) + TimeSpan.FromHours(23);

    public static IEndpointRouteBuilder MapWebhookEndpoints(this IEndpointRouteBuilder app)
    {
        app.MapPost("/api/webhooks/graph", async (
            [FromQuery] string? validationToken,
            HttpRequest request,
            ILoggerFactory loggerFactory,
            CancellationToken cancellationToken) =>
        {
            if (!string.IsNullOrEmpty(validationToken))
            {
                return Results.Text(validationToken, "text/plain");
            }

            var logger = loggerFactory.CreateLogger("GraphWebhooks");
            var payload = await JsonSerializer.DeserializeAsync<GraphWebhookPayload>(
                request.Body,
                new JsonSerializerOptions(JsonSerializerDefaults.Web),
                cancellationToken);

            foreach (var notification in payload?.Value ?? [])
            {
                logger.LogInformation(
                    "Received Graph webhook notification {SubscriptionId} {ChangeType} {Resource} {ResourceId} {TenantId}",
                    notification.SubscriptionId,
                    notification.ChangeType,
                    notification.Resource,
                    notification.ResourceData?.Id,
                    notification.TenantId);
            }

            return Results.Accepted();
        }).AllowAnonymous().WithTags("Webhooks");

        var subscriptions = app.MapGroup("/api/webhooks/subscriptions").RequireAuthorization().WithTags("Webhooks");

        subscriptions.MapPost(string.Empty, async (
            [FromBody] CreateSubscriptionRequest request,
            IAppOnlyGraphClient appOnlyGraphClient,
            IGraphSubscriptionStore subscriptionStore,
            IOptions<WebhookOptions> options,
            CancellationToken cancellationToken) =>
        {
            if (string.IsNullOrWhiteSpace(request.DriveId))
            {
                return EndpointResults.BadRequestProblem("Drive ID is required", "Provide a driveId to subscribe to.");
            }

            var expiration = DateTimeOffset.UtcNow.Add(SubscriptionLifetime);
            var resource = $"/drives/{request.DriveId}/root";
            var subscription = await appOnlyGraphClient.Client.Subscriptions.PostAsync(new Subscription
            {
                ChangeType = "updated",
                NotificationUrl = options.Value.NotificationUrl,
                Resource = resource,
                ExpirationDateTime = expiration,
                ClientState = options.Value.ClientStateSecret
            }, cancellationToken: cancellationToken);

            if (subscription?.Id is null)
            {
                return EndpointResults.NotFoundProblem("Subscription creation failed", "Microsoft Graph did not return a subscription ID.");
            }

            var record = new GraphSubscriptionRecord(
                subscription.Id,
                request.DriveId,
                subscription.Resource ?? resource,
                subscription.ExpirationDateTime ?? expiration,
                subscription.ClientState ?? options.Value.ClientStateSecret);

            subscriptionStore.Upsert(record);
            return Results.Ok(record);
        });

        subscriptions.MapGet(string.Empty, (IGraphSubscriptionStore subscriptionStore) => Results.Ok(subscriptionStore.List()));

        subscriptions.MapDelete("/{id}", async (
            [FromRoute] string id,
            IAppOnlyGraphClient appOnlyGraphClient,
            IGraphSubscriptionStore subscriptionStore,
            CancellationToken cancellationToken) =>
        {
            await appOnlyGraphClient.Client.Subscriptions[id].DeleteAsync(cancellationToken: cancellationToken);
            subscriptionStore.Remove(id, out _);
            return Results.NoContent();
        });

        return app;
    }
}
