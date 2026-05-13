using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using Wordembedded.Api.Contracts;

namespace Wordembedded.Api.Endpoints;

public static class RecycleBinEndpoints
{
    private const string Scope = "https://graph.microsoft.com/FileStorageContainer.Selected";

    public static IEndpointRouteBuilder MapRecycleBinEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api/containers/{containerId}/recyclebin").RequireAuthorization().WithTags("Recycle bin");

        group.MapGet(string.Empty, async (
            [FromRoute] string containerId,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            var token = await tokenAcquisition.GetAccessTokenForUserAsync(new[] { Scope });
            var client = httpClientFactory.CreateClient("graph-recyclebin");
            var request = new HttpRequestMessage(HttpMethod.Get,
                $"https://graph.microsoft.com/v1.0/storage/fileStorage/containers/{Uri.EscapeDataString(containerId)}/recycleBin/items");
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var response = await client.SendAsync(request, cancellationToken);
            var raw = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                return Results.Problem(title: "Could not list recycle bin", detail: raw, statusCode: (int)response.StatusCode);
            }

            using var doc = JsonDocument.Parse(raw);
            var items = new List<RecycleBinItemResponse>();
            if (doc.RootElement.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
            {
                foreach (var element in value.EnumerateArray())
                {
                    items.Add(new RecycleBinItemResponse(
                        TryGetString(element, "id"),
                        TryGetString(element, "name"),
                        TryGetLong(element, "size"),
                        TryGetDateTimeOffset(element, "deletedDateTime"),
                        ExtractIdentity(element, "deletedBy"),
                        TryGetDateTimeOffset(element, "lastModifiedDateTime"),
                        TryGetString(element, "originalDriveItemId")));
                }
            }

            return Results.Ok(new { value = items });
        });

        group.MapPost("/restore", async (
            [FromRoute] string containerId,
            [FromBody] RecycleBinActionRequest request,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            if (request?.Ids is null || request.Ids.Count == 0)
            {
                return EndpointResults.BadRequestProblem("ids is required", "Provide at least one recycle bin item id.");
            }

            var token = await tokenAcquisition.GetAccessTokenForUserAsync(new[] { Scope });
            var client = httpClientFactory.CreateClient("graph-recyclebin");
            var graphReq = new HttpRequestMessage(HttpMethod.Post,
                $"https://graph.microsoft.com/v1.0/storage/fileStorage/containers/{Uri.EscapeDataString(containerId)}/recycleBin/items/restore")
            {
                Content = JsonContent.Create(new { ids = request.Ids })
            };
            graphReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var response = await client.SendAsync(graphReq, cancellationToken);
            var raw = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                return Results.Problem(title: "Restore failed", detail: raw, statusCode: (int)response.StatusCode);
            }

            return Results.Ok(new { restored = request.Ids });
        });

        group.MapDelete("/items", async (
            [FromRoute] string containerId,
            [FromBody] RecycleBinActionRequest request,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            if (request?.Ids is null || request.Ids.Count == 0)
            {
                return EndpointResults.BadRequestProblem("ids is required", "Provide at least one recycle bin item id.");
            }

            var token = await tokenAcquisition.GetAccessTokenForUserAsync(new[] { Scope });
            var client = httpClientFactory.CreateClient("graph-recyclebin");
            var graphReq = new HttpRequestMessage(HttpMethod.Delete,
                $"https://graph.microsoft.com/v1.0/storage/fileStorage/containers/{Uri.EscapeDataString(containerId)}/recycleBin/items")
            {
                Content = JsonContent.Create(new { ids = request.Ids })
            };
            graphReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var response = await client.SendAsync(graphReq, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                var raw = await response.Content.ReadAsStringAsync(cancellationToken);
                return Results.Problem(title: "Permanent delete failed", detail: raw, statusCode: (int)response.StatusCode);
            }

            return Results.NoContent();
        });

        return app;
    }

    private static string? TryGetString(JsonElement element, string property)
        => element.TryGetProperty(property, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static long? TryGetLong(JsonElement element, string property)
    {
        if (!element.TryGetProperty(property, out var v)) return null;
        return v.ValueKind == JsonValueKind.Number && v.TryGetInt64(out var l) ? l : null;
    }

    private static DateTimeOffset? TryGetDateTimeOffset(JsonElement element, string property)
    {
        var s = TryGetString(element, property);
        return string.IsNullOrWhiteSpace(s) ? null : DateTimeOffset.Parse(s);
    }

    private static IdentityResponse? ExtractIdentity(JsonElement element, string property)
    {
        if (!element.TryGetProperty(property, out var idSet)) return null;
        var user = idSet.TryGetProperty("user", out var u) ? u : default;
        if (user.ValueKind == JsonValueKind.Object)
        {
            return new IdentityResponse(
                TryGetString(user, "displayName"),
                TryGetString(user, "email"));
        }
        return null;
    }
}
