using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace Wordembedded.Api.Services;

public sealed class WebhookRenewalService(
    IAppOnlyGraphClient appOnlyGraphClient,
    IGraphSubscriptionStore subscriptionStore,
    ILogger<WebhookRenewalService> logger) : BackgroundService
{
    private static readonly TimeSpan RenewalInterval = TimeSpan.FromHours(12);
    private static readonly TimeSpan SubscriptionLifetime = TimeSpan.FromDays(2) + TimeSpan.FromHours(23);

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        using var timer = new PeriodicTimer(RenewalInterval);

        while (!stoppingToken.IsCancellationRequested)
        {
            await RenewSubscriptionsAsync(stoppingToken);

            try
            {
                await timer.WaitForNextTickAsync(stoppingToken);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }
        }
    }

    private async Task RenewSubscriptionsAsync(CancellationToken cancellationToken)
    {
        foreach (var record in subscriptionStore.List())
        {
            var newExpiration = DateTimeOffset.UtcNow.Add(SubscriptionLifetime);

            try
            {
                var updated = await appOnlyGraphClient.Client.Subscriptions[record.Id].PatchAsync(
                    new Subscription { ExpirationDateTime = newExpiration },
                    cancellationToken: cancellationToken);

                subscriptionStore.Upsert(record with { ExpirationDateTime = updated?.ExpirationDateTime ?? newExpiration });
                logger.LogInformation("Renewed Graph subscription {SubscriptionId} for drive {DriveId} until {ExpirationDateTime}", record.Id, record.DriveId, newExpiration);
            }
            catch (ODataError ex) when (ex.ResponseStatusCode == StatusCodes.Status404NotFound)
            {
                subscriptionStore.Remove(record.Id, out _);
                logger.LogWarning("Dropped missing Graph subscription {SubscriptionId} for drive {DriveId}", record.Id, record.DriveId);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                logger.LogError(ex, "Failed to renew Graph subscription {SubscriptionId} for drive {DriveId}", record.Id, record.DriveId);
            }
        }
    }
}
