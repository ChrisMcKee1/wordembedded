using System.Net.Http.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using Wordembedded.Api.Contracts;

namespace Wordembedded.Api.Endpoints;

public static class PreviewEndpoints
{
    private const string PreviewScope = "https://graph.microsoft.com/FileStorageContainer.Selected";

    public static IEndpointRouteBuilder MapPreviewEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api/drives/{driveId}/items/{itemId}/preview").RequireAuthorization().WithTags("Preview");

        // Bypass the Graph SDK and call /preview directly: the SDK's serializer omits an empty
        // body when both Page and Zoom are null, and Graph rejects an empty body with 400
        // invalidRequest. A literal "{}" body is what the docs specify.
        group.MapPost(string.Empty, async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            [FromBody] PreviewRequest? request,
            HttpResponse response,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            ILogger<Marker> logger,
            CancellationToken cancellationToken) =>
        {
            var token = await tokenAcquisition.GetAccessTokenForUserAsync(new[] { PreviewScope });
            var client = httpClientFactory.CreateClient("graph-preview");
            var payload = new Dictionary<string, object?>();
            if (!string.IsNullOrWhiteSpace(request?.Page)) payload["page"] = request.Page;
            if (request?.Zoom is { } zoom) payload["zoom"] = zoom;
            if (!string.IsNullOrWhiteSpace(request?.Viewer)) payload["viewer"] = request.Viewer;
            if (request?.AllowEdit is { } allowEdit) payload["allowEdit"] = allowEdit;
            if (request?.Chromeless is { } chromeless) payload["chromeless"] = chromeless;

            // Beta endpoint supports allowEdit/viewer/chromeless. v1.0 only honors page/zoom.
            var useBeta = request?.AllowEdit == true
                || !string.IsNullOrWhiteSpace(request?.Viewer)
                || request?.Chromeless is not null;
            var graphVersion = useBeta ? "beta" : "v1.0";

            var httpRequest = new HttpRequestMessage(HttpMethod.Post,
                $"https://graph.microsoft.com/{graphVersion}/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(itemId)}/preview")
            {
                Content = JsonContent.Create(payload)
            };
            httpRequest.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var graphResponse = await client.SendAsync(httpRequest, cancellationToken);
            var body = await graphResponse.Content.ReadAsStringAsync(cancellationToken);

            if (!graphResponse.IsSuccessStatusCode)
            {
                logger.LogWarning("Preview Graph call failed {Status}: {Body}", (int)graphResponse.StatusCode, body);
                return Results.Problem(
                    title: "Preview unavailable",
                    detail: body,
                    statusCode: (int)graphResponse.StatusCode);
            }

            var info = System.Text.Json.JsonSerializer.Deserialize<PreviewInfo>(body, new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            response.Headers.CacheControl = "private, max-age=30";

            return info is null
                ? EndpointResults.NotFoundProblem("Preview unavailable", "Microsoft Graph did not return preview information.")
                : Results.Ok(new PreviewResponse(AppendNoBanner(info.GetUrl), info.PostUrl, info.PostParameters));
        });

        return app;
    }

    private static string? AppendNoBanner(string? getUrl)
    {
        if (string.IsNullOrWhiteSpace(getUrl) || getUrl.Contains("nb=true", StringComparison.OrdinalIgnoreCase))
        {
            return getUrl;
        }

        return getUrl.Contains('?', StringComparison.Ordinal) ? $"{getUrl}&nb=true" : $"{getUrl}?nb=true";
    }

    public sealed class Marker { }

    private sealed record PreviewInfo(string? GetUrl, string? PostUrl, string? PostParameters);
}
