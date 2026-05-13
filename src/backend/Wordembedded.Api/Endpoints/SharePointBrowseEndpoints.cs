using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using Wordembedded.Api.Contracts;

namespace Wordembedded.Api.Endpoints;

public static class SharePointBrowseEndpoints
{
    public static IEndpointRouteBuilder MapSharePointBrowseEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api/sharepoint").RequireAuthorization().WithTags("SharePoint browse");

        // Microsoft Search for SharePoint driveItems (Word/Excel/PowerPoint/PDF/etc.)
        group.MapGet("/search", async (
            [FromQuery(Name = "q")] string query,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return EndpointResults.BadRequestProblem("Query required", "Pass a non-empty q parameter.");
            }

            var token = await tokenAcquisition.GetAccessTokenForUserAsync([
                "https://graph.microsoft.com/Sites.Read.All",
                "https://graph.microsoft.com/Files.Read.All",
            ]);
            var client = httpClientFactory.CreateClient("graph-sp-browse");
            var body = new
            {
                requests = new[] {
                    new {
                        entityTypes = new[] { "driveItem" },
                        query = new { queryString = query },
                        from = 0,
                        size = 25,
                    }
                }
            };

            var req = new HttpRequestMessage(HttpMethod.Post, "https://graph.microsoft.com/v1.0/search/query")
            {
                Content = JsonContent.Create(body)
            };
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var resp = await client.SendAsync(req, cancellationToken);
            var raw = await resp.Content.ReadAsStringAsync(cancellationToken);
            if (!resp.IsSuccessStatusCode)
            {
                return Results.Problem(title: "SharePoint search failed", detail: raw, statusCode: (int)resp.StatusCode);
            }

            using var doc = JsonDocument.Parse(raw);
            var hits = new List<SharePointSearchHit>();
            if (doc.RootElement.TryGetProperty("value", out var responses) && responses.ValueKind == JsonValueKind.Array)
            {
                foreach (var response in responses.EnumerateArray())
                {
                    if (!response.TryGetProperty("hitsContainers", out var hitsContainers) || hitsContainers.ValueKind != JsonValueKind.Array) continue;
                    foreach (var container in hitsContainers.EnumerateArray())
                    {
                        if (!container.TryGetProperty("hits", out var rawHits) || rawHits.ValueKind != JsonValueKind.Array) continue;
                        foreach (var hit in rawHits.EnumerateArray())
                        {
                            if (!hit.TryGetProperty("resource", out var resource)) continue;
                            var resourceType = resource.TryGetProperty("@odata.type", out var typeEl) ? typeEl.GetString() : null;
                            if (resourceType is not null && !resourceType.Contains("driveItem", StringComparison.OrdinalIgnoreCase)) continue;

                            var parent = resource.TryGetProperty("parentReference", out var pr) ? pr : default;
                            hits.Add(new SharePointSearchHit(
                                hit.TryGetProperty("hitId", out var hitId) ? hitId.GetString() : null,
                                resource.TryGetProperty("id", out var id) ? id.GetString() : null,
                                resource.TryGetProperty("name", out var name) ? name.GetString() : null,
                                resource.TryGetProperty("webUrl", out var webUrl) ? webUrl.GetString() : null,
                                resource.TryGetProperty("size", out var size) && size.ValueKind == JsonValueKind.Number ? size.GetInt64() : null,
                                resource.TryGetProperty("lastModifiedDateTime", out var mod) ? mod.GetString() : null,
                                parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty("driveId", out var driveId) ? driveId.GetString() : null,
                                parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty("path", out var path) ? path.GetString() : null,
                                parent.ValueKind == JsonValueKind.Object && parent.TryGetProperty("siteId", out var siteId) ? siteId.GetString() : null,
                                hit.TryGetProperty("summary", out var summary) ? summary.GetString() : null
                            ));
                        }
                    }
                }
            }

            return Results.Ok(new { value = hits });
        });

        // List SharePoint sites the user can access. Surface a flat list of the most relevant sites
        // returned by /sites?search=*.
        group.MapGet("/sites", async (
            [FromQuery] string? q,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            var token = await tokenAcquisition.GetAccessTokenForUserAsync(["https://graph.microsoft.com/Sites.Read.All"]);
            var client = httpClientFactory.CreateClient("graph-sp-browse");
            var search = string.IsNullOrWhiteSpace(q) ? "*" : q;
            var req = new HttpRequestMessage(HttpMethod.Get,
                $"https://graph.microsoft.com/v1.0/sites?search={Uri.EscapeDataString(search)}");
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var resp = await client.SendAsync(req, cancellationToken);
            var raw = await resp.Content.ReadAsStringAsync(cancellationToken);
            if (!resp.IsSuccessStatusCode)
            {
                return Results.Problem(title: "Could not list sites", detail: raw, statusCode: (int)resp.StatusCode);
            }
            using var doc = JsonDocument.Parse(raw);
            var sites = new List<SharePointSiteSummary>();
            if (doc.RootElement.TryGetProperty("value", out var sitesArr) && sitesArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var s in sitesArr.EnumerateArray())
                {
                    sites.Add(new SharePointSiteSummary(
                        s.TryGetProperty("id", out var id) ? id.GetString() : null,
                        s.TryGetProperty("displayName", out var dn) ? dn.GetString() : null,
                        s.TryGetProperty("name", out var n) ? n.GetString() : null,
                        s.TryGetProperty("webUrl", out var w) ? w.GetString() : null
                    ));
                }
            }
            return Results.Ok(new { value = sites });
        });

        // List document libraries (drives) on a SharePoint site.
        group.MapGet("/sites/{siteId}/drives", async (
            [FromRoute] string siteId,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            var token = await tokenAcquisition.GetAccessTokenForUserAsync(["https://graph.microsoft.com/Sites.Read.All"]);
            var client = httpClientFactory.CreateClient("graph-sp-browse");
            var req = new HttpRequestMessage(HttpMethod.Get,
                $"https://graph.microsoft.com/v1.0/sites/{Uri.EscapeDataString(siteId)}/drives");
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var resp = await client.SendAsync(req, cancellationToken);
            var raw = await resp.Content.ReadAsStringAsync(cancellationToken);
            if (!resp.IsSuccessStatusCode)
            {
                return Results.Problem(title: "Could not list drives", detail: raw, statusCode: (int)resp.StatusCode);
            }
            using var doc = JsonDocument.Parse(raw);
            var drives = new List<SharePointDriveSummary>();
            if (doc.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var d in arr.EnumerateArray())
                {
                    drives.Add(new SharePointDriveSummary(
                        d.TryGetProperty("id", out var id) ? id.GetString() : null,
                        d.TryGetProperty("name", out var n) ? n.GetString() : null,
                        d.TryGetProperty("webUrl", out var w) ? w.GetString() : null,
                        d.TryGetProperty("driveType", out var dt) ? dt.GetString() : null
                    ));
                }
            }
            return Results.Ok(new { value = drives });
        });

        // List children of a folder in a SharePoint drive (root if itemId omitted).
        group.MapGet("/drives/{driveId}/children", async (
            [FromRoute] string driveId,
            [FromQuery] string? itemId,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            var token = await tokenAcquisition.GetAccessTokenForUserAsync(["https://graph.microsoft.com/Files.Read.All"]);
            var client = httpClientFactory.CreateClient("graph-sp-browse");
            var path = string.IsNullOrWhiteSpace(itemId)
                ? $"https://graph.microsoft.com/v1.0/drives/{Uri.EscapeDataString(driveId)}/root/children"
                : $"https://graph.microsoft.com/v1.0/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(itemId)}/children";
            var req = new HttpRequestMessage(HttpMethod.Get, path);
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            using var resp = await client.SendAsync(req, cancellationToken);
            var raw = await resp.Content.ReadAsStringAsync(cancellationToken);
            if (!resp.IsSuccessStatusCode)
            {
                return Results.Problem(title: "Could not list children", detail: raw, statusCode: (int)resp.StatusCode);
            }
            using var doc = JsonDocument.Parse(raw);
            var items = new List<SharePointBrowseItem>();
            if (doc.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                foreach (var i in arr.EnumerateArray())
                {
                    items.Add(new SharePointBrowseItem(
                        i.TryGetProperty("id", out var id) ? id.GetString() : null,
                        i.TryGetProperty("name", out var n) ? n.GetString() : null,
                        i.TryGetProperty("webUrl", out var w) ? w.GetString() : null,
                        i.TryGetProperty("size", out var sz) && sz.ValueKind == JsonValueKind.Number ? sz.GetInt64() : null,
                        i.TryGetProperty("lastModifiedDateTime", out var mod) ? mod.GetString() : null,
                        i.TryGetProperty("folder", out _),
                        i.TryGetProperty("file", out var f) && f.TryGetProperty("mimeType", out var mt) ? mt.GetString() : null
                    ));
                }
            }
            return Results.Ok(new { value = items });
        });

        return app;
    }
}
