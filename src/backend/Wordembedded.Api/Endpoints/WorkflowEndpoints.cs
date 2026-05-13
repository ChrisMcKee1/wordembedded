using System.Net.Http.Json;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using Wordembedded.Api.Contracts;

namespace Wordembedded.Api.Endpoints;

public static class WorkflowEndpoints
{
    private static readonly string[] FileStorageScopes = ["https://graph.microsoft.com/FileStorageContainer.Selected"];
    private const string StatusPrefix = "[wf]";

    public static IEndpointRouteBuilder MapWorkflowEndpoints(this IEndpointRouteBuilder app)
    {
        var group = app.MapGroup("/api").RequireAuthorization().WithTags("Workflow");

        // Cross-drive copy: pull a SharePoint driveItem into our SPE container.
        // Uses Graph copy API (POST /drives/{src}/items/{id}/copy) with parentReference pointing at the destination drive.
        // Returns Location header to monitor async progress.
        group.MapPost("/sharepoint/import", async (
            [FromBody] ImportFileRequest request,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            if (string.IsNullOrWhiteSpace(request?.SourceDriveId) ||
                string.IsNullOrWhiteSpace(request.SourceItemId) ||
                string.IsNullOrWhiteSpace(request.TargetDriveId))
            {
                return EndpointResults.BadRequestProblem("Missing source or target", "sourceDriveId, sourceItemId, and targetDriveId are required.");
            }

            var token = await tokenAcquisition.GetAccessTokenForUserAsync(["https://graph.microsoft.com/FileStorageContainer.Selected", "https://graph.microsoft.com/Files.Read.All"]);
            var client = httpClientFactory.CreateClient("graph-workflow");

            // conflictBehavior is a query parameter (not body) per Microsoft Graph docs. Default to
            // "rename" so a re-import won't silently clobber a file in the container.
            var copyUri = $"https://graph.microsoft.com/v1.0/drives/{Uri.EscapeDataString(request.SourceDriveId)}/items/{Uri.EscapeDataString(request.SourceItemId)}/copy?@microsoft.graph.conflictBehavior=rename";

            var body = new Dictionary<string, object?>
            {
                ["parentReference"] = new Dictionary<string, object?>
                {
                    ["driveId"] = request.TargetDriveId,
                    ["id"] = string.IsNullOrWhiteSpace(request.TargetFolderId) ? "root" : request.TargetFolderId,
                },
            };
            if (!string.IsNullOrWhiteSpace(request.NewName))
            {
                body["name"] = request.NewName;
            }

            var graphReq = new HttpRequestMessage(HttpMethod.Post, copyUri)
            {
                Content = JsonContent.Create(body)
            };
            graphReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var response = await client.SendAsync(graphReq, cancellationToken);
            var raw = await response.Content.ReadAsStringAsync(cancellationToken);
            if (response.StatusCode != System.Net.HttpStatusCode.Accepted && !response.IsSuccessStatusCode)
            {
                return Results.Problem(title: "Import failed", detail: raw, statusCode: (int)response.StatusCode);
            }

            var monitorUrl = response.Headers.Location?.ToString();
            return Results.Accepted(monitorUrl, new ImportResponse(monitorUrl, "accepted"));
        });

        // Cross-drive copy: publish from SPE container to a SharePoint destination.
        // Sets workflow status on the source item to Published once the copy is accepted.
        group.MapPost("/drives/{driveId}/items/{itemId}/publish", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            [FromBody] PublishRequest request,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            ILogger<WorkflowLogger> logger,
            CancellationToken cancellationToken) =>
        {
            if (string.IsNullOrWhiteSpace(request?.TargetDriveId))
            {
                return EndpointResults.BadRequestProblem("Missing target", "targetDriveId is required.");
            }

            var token = await tokenAcquisition.GetAccessTokenForUserAsync(["https://graph.microsoft.com/FileStorageContainer.Selected", "https://graph.microsoft.com/Sites.ReadWrite.All"]);
            var client = httpClientFactory.CreateClient("graph-workflow");

            // Per Microsoft Graph docs, `@microsoft.graph.conflictBehavior` is a QUERY parameter
            // (not body) and supports "fail" (default), "replace" (overwrite, deletes destination's
            // version history), "rename" (append a unique integer). `replace` is supported for file
            // items; folder conflicts always behave like `fail`. Docs:
            // https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0
            //
            // Known issue from the docs: `includeAllVersionHistory` is ignored if `name` is also
            // set. We omit `name` from the body by default so the destination keeps the source's
            // filename and version history is preserved.
            var conflictBehavior = string.IsNullOrWhiteSpace(request.ConflictBehavior) ? "replace" : request.ConflictBehavior;
            var copyUri = $"https://graph.microsoft.com/v1.0/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(itemId)}/copy"
                + $"?@microsoft.graph.conflictBehavior={Uri.EscapeDataString(conflictBehavior)}";

            var body = new Dictionary<string, object?>
            {
                ["parentReference"] = new Dictionary<string, object?>
                {
                    ["driveId"] = request.TargetDriveId,
                    ["id"] = string.IsNullOrWhiteSpace(request.TargetFolderId) ? "root" : request.TargetFolderId,
                },
                ["includeAllVersionHistory"] = true,
            };
            if (!string.IsNullOrWhiteSpace(request.NewName))
            {
                body["name"] = request.NewName;
            }

            var graphReq = new HttpRequestMessage(HttpMethod.Post, copyUri)
            {
                Content = JsonContent.Create(body)
            };
            graphReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var response = await client.SendAsync(graphReq, cancellationToken);
            var raw = await response.Content.ReadAsStringAsync(cancellationToken);
            if (response.StatusCode != System.Net.HttpStatusCode.Accepted && !response.IsSuccessStatusCode)
            {
                logger.LogWarning("Publish copy failed {Status}: {Body}", (int)response.StatusCode, raw);
                return Results.Problem(title: "Publish failed", detail: raw, statusCode: (int)response.StatusCode);
            }
            var monitorUrl = response.Headers.Location?.ToString();

            // Poll the monitor URL until the async copy completes so we can report the truth
            // instead of optimistically marking the file as Published.
            var (finalStatus, finalDetail, resourceId) = await PollMonitorAsync(client, monitorUrl, cancellationToken);
            logger.LogInformation("Publish poll terminal status: {Status}; resourceId: {Resource}; detail: {Detail}", finalStatus, resourceId, finalDetail);

            if (string.Equals(finalStatus, "completed", StringComparison.OrdinalIgnoreCase) || string.Equals(finalStatus, "succeeded", StringComparison.OrdinalIgnoreCase))
            {
                // Look up the destination item so we can carry the webUrl forward in the source status.
                PublishedReference? destination = null;
                if (!string.IsNullOrWhiteSpace(resourceId))
                {
                    try
                    {
                        var lookupReq = new HttpRequestMessage(HttpMethod.Get,
                            $"https://graph.microsoft.com/v1.0/drives/{Uri.EscapeDataString(request.TargetDriveId)}/items/{Uri.EscapeDataString(resourceId)}?$select=id,name,webUrl");
                        lookupReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                        using var lookup = await client.SendAsync(lookupReq, cancellationToken);
                        if (lookup.IsSuccessStatusCode)
                        {
                            var lookupBody = await lookup.Content.ReadAsStringAsync(cancellationToken);
                            using var lookupDoc = JsonDocument.Parse(lookupBody);
                            destination = new PublishedReference(
                                request.TargetDriveId,
                                resourceId,
                                lookupDoc.RootElement.TryGetProperty("webUrl", out var w) ? w.GetString() : null,
                                lookupDoc.RootElement.TryGetProperty("name", out var n) ? n.GetString() : null,
                                DateTimeOffset.UtcNow);
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.LogWarning(ex, "Could not resolve destination item {ResourceId}", resourceId);
                    }
                }

                await SetStatusAsync(client, token, driveId, itemId, WorkflowStatus.Published, cancellationToken, destination);
                return Results.Ok(new PublishResponse(monitorUrl, "Published", destination));
            }

            return Results.Problem(
                title: "Publish copy did not complete",
                detail: finalDetail ?? $"Final status was '{finalStatus}'. Source file was NOT marked Published.",
                statusCode: StatusCodes.Status502BadGateway);
        });

        // Get the current workflow status of a driveItem.
        group.MapGet("/drives/{driveId}/items/{itemId}/workflow", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            var token = await tokenAcquisition.GetAccessTokenForUserAsync(FileStorageScopes);
            var client = httpClientFactory.CreateClient("graph-workflow");
            var graphReq = new HttpRequestMessage(HttpMethod.Get,
                $"https://graph.microsoft.com/v1.0/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(itemId)}?$select=description,name");
            graphReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var response = await client.SendAsync(graphReq, cancellationToken);
            var raw = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                return Results.Problem(title: "Could not read item", detail: raw, statusCode: (int)response.StatusCode);
            }

            using var doc = JsonDocument.Parse(raw);
            var description = doc.RootElement.TryGetProperty("description", out var d) ? d.GetString() : null;
            var status = ExtractStatus(description) ?? WorkflowStatus.Draft;
            var publishedReference = ExtractPublishedReference(description);
            return Results.Ok(new WorkflowStatusResponse(status.ToString(), publishedReference));
        });

        // Set the workflow status of a driveItem.
        group.MapPut("/drives/{driveId}/items/{itemId}/workflow", async (
            [FromRoute] string driveId,
            [FromRoute] string itemId,
            [FromBody] SetWorkflowRequest request,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            if (!Enum.TryParse<WorkflowStatus>(request.Status, ignoreCase: true, out var status))
            {
                return EndpointResults.BadRequestProblem("Invalid status", "Status must be Draft, InReview, Approved, or Published.");
            }

            var token = await tokenAcquisition.GetAccessTokenForUserAsync(["https://graph.microsoft.com/FileStorageContainer.Selected", "https://graph.microsoft.com/Files.Read.All"]);
            var client = httpClientFactory.CreateClient("graph-workflow");
            await SetStatusAsync(client, token, driveId, itemId, status, cancellationToken);
            return Results.Ok(new WorkflowStatusResponse(status.ToString(), null));
        });

        // Configure the publish target stored on the container's customProperties.
        group.MapGet("/containers/{containerId}/publish-target", async (
            [FromRoute] string containerId,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            var token = await tokenAcquisition.GetAccessTokenForUserAsync(FileStorageScopes);
            var client = httpClientFactory.CreateClient("graph-workflow");
            var graphReq = new HttpRequestMessage(HttpMethod.Get,
                $"https://graph.microsoft.com/v1.0/storage/fileStorage/containers/{Uri.EscapeDataString(containerId)}/customProperties");
            graphReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var response = await client.SendAsync(graphReq, cancellationToken);
            var raw = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                // 404 = no custom properties set
                if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    return Results.Ok(new PublishTargetResponse(null, null, null, null));
                }
                return Results.Problem(title: "Could not read publish target", detail: raw, statusCode: (int)response.StatusCode);
            }

            using var doc = JsonDocument.Parse(raw);
            var target = doc.RootElement.TryGetProperty("publishTarget", out var pt) && pt.TryGetProperty("value", out var pv)
                ? pv.GetString()
                : null;
            if (string.IsNullOrWhiteSpace(target))
            {
                return Results.Ok(new PublishTargetResponse(null, null, null, null));
            }

            try
            {
                var parsed = JsonSerializer.Deserialize<PublishTargetResponse>(target!, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                return Results.Ok(parsed ?? new PublishTargetResponse(null, null, null, null));
            }
            catch
            {
                return Results.Ok(new PublishTargetResponse(null, null, null, null));
            }
        });

        group.MapPut("/containers/{containerId}/publish-target", async (
            [FromRoute] string containerId,
            [FromBody] PublishTargetResponse request,
            ITokenAcquisition tokenAcquisition,
            IHttpClientFactory httpClientFactory,
            CancellationToken cancellationToken) =>
        {
            var token = await tokenAcquisition.GetAccessTokenForUserAsync(FileStorageScopes);
            var client = httpClientFactory.CreateClient("graph-workflow");
            var payload = new Dictionary<string, object?>
            {
                ["publishTarget"] = new Dictionary<string, object?>
                {
                    ["value"] = JsonSerializer.Serialize(request),
                    ["isSearchable"] = false,
                }
            };

            var graphReq = new HttpRequestMessage(HttpMethod.Patch,
                $"https://graph.microsoft.com/v1.0/storage/fileStorage/containers/{Uri.EscapeDataString(containerId)}/customProperties")
            {
                Content = JsonContent.Create(payload)
            };
            graphReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            using var response = await client.SendAsync(graphReq, cancellationToken);
            var raw = await response.Content.ReadAsStringAsync(cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                return Results.Problem(title: "Could not save publish target", detail: raw, statusCode: (int)response.StatusCode);
            }

            return Results.Ok(request);
        });

        return app;
    }

    private static async Task<PublishedReference?> GetPublishedReferenceAsync(HttpClient client, string token, string driveId, string itemId, CancellationToken cancellationToken)
    {
        var req = new HttpRequestMessage(HttpMethod.Get,
            $"https://graph.microsoft.com/v1.0/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(itemId)}?$select=description");
        req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
        using var resp = await client.SendAsync(req, cancellationToken);
        if (!resp.IsSuccessStatusCode) return null;
        var body = await resp.Content.ReadAsStringAsync(cancellationToken);
        using var doc = JsonDocument.Parse(body);
        if (!doc.RootElement.TryGetProperty("description", out var d)) return null;
        return ExtractPublishedReference(d.GetString());
    }

    private static async Task SetStatusAsync(HttpClient client, string token, string driveId, string itemId, WorkflowStatus status, CancellationToken cancellationToken, PublishedReference? publishedRef = null)
    {
        // Fetch existing description so we can preserve any user description below the status line.
        var getReq = new HttpRequestMessage(HttpMethod.Get,
            $"https://graph.microsoft.com/v1.0/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(itemId)}?$select=description");
        getReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
        using var current = await client.SendAsync(getReq, cancellationToken);
        var currentBody = await current.Content.ReadAsStringAsync(cancellationToken);
        string? userDescription = null;
        PublishedReference? existingRef = null;
        if (current.IsSuccessStatusCode)
        {
            using var doc = JsonDocument.Parse(currentBody);
            if (doc.RootElement.TryGetProperty("description", out var d))
            {
                var existing = d.GetString();
                userDescription = StripStatus(existing);
                existingRef = ExtractPublishedReference(existing);
            }
        }

        // Carry forward an existing published reference unless this call is providing a new one.
        // Clear it when transitioning to states other than Published.
        var refToWrite = publishedRef ?? (status == WorkflowStatus.Published ? existingRef : null);

        var statusLine = $"{StatusPrefix} {status}";
        if (refToWrite is not null)
        {
            // Embed published-reference JSON on the same line. Format: "[wf] Published {\"...\"}"
            var refJson = JsonSerializer.Serialize(refToWrite, new JsonSerializerOptions { DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull });
            statusLine = $"{statusLine} {refJson}";
        }

        var description = string.IsNullOrWhiteSpace(userDescription)
            ? statusLine
            : $"{statusLine}\n{userDescription}";

        var patchReq = new HttpRequestMessage(HttpMethod.Patch,
            $"https://graph.microsoft.com/v1.0/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(itemId)}")
        {
            Content = JsonContent.Create(new { description })
        };
        patchReq.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
        using var patchResponse = await client.SendAsync(patchReq, cancellationToken);
        if (!patchResponse.IsSuccessStatusCode)
        {
            var errBody = await patchResponse.Content.ReadAsStringAsync(cancellationToken);
            throw new InvalidOperationException($"Failed to set status: {(int)patchResponse.StatusCode} {errBody}");
        }
    }

    private static WorkflowStatus? ExtractStatus(string? description)
    {
        if (string.IsNullOrWhiteSpace(description)) return null;
        var firstLine = description.Split('\n').FirstOrDefault()?.Trim();
        if (string.IsNullOrWhiteSpace(firstLine) || !firstLine.StartsWith(StatusPrefix, StringComparison.OrdinalIgnoreCase)) return null;
        var rest = firstLine[StatusPrefix.Length..].Trim();
        // Strip any embedded JSON suffix so "Published {\"...\"}" still parses to Published.
        var braceIdx = rest.IndexOf('{');
        if (braceIdx > 0) rest = rest[..braceIdx].Trim();
        return Enum.TryParse<WorkflowStatus>(rest, ignoreCase: true, out var status) ? status : null;
    }

    private static PublishedReference? ExtractPublishedReference(string? description)
    {
        if (string.IsNullOrWhiteSpace(description)) return null;
        var firstLine = description.Split('\n').FirstOrDefault()?.Trim();
        if (string.IsNullOrWhiteSpace(firstLine)) return null;
        var braceIdx = firstLine.IndexOf('{');
        if (braceIdx < 0) return null;
        var jsonPart = firstLine[braceIdx..];
        try
        {
            return JsonSerializer.Deserialize<PublishedReference>(jsonPart, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch
        {
            return null;
        }
    }

    private static string? StripStatus(string? description)
    {
        if (string.IsNullOrWhiteSpace(description)) return null;
        var lines = description.Split('\n').ToList();
        if (lines.Count == 0) return description;
        if (lines[0].TrimStart().StartsWith(StatusPrefix, StringComparison.OrdinalIgnoreCase))
        {
            lines.RemoveAt(0);
        }
        return string.Join('\n', lines).Trim();
    }

    public sealed class WorkflowLogger { }

    // Poll a Graph async-monitor URL until terminal state. Returns (status, error, resourceId).
    private static async Task<(string Status, string? Error, string? ResourceId)> PollMonitorAsync(HttpClient client, string? monitorUrl, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(monitorUrl))
        {
            return ("unknown", "Graph did not return a Location header for the async operation.", null);
        }

        // Up to ~30s of polling at increasing intervals.
        var delays = new[] { 500, 1000, 2000, 3000, 5000, 5000, 5000, 5000 };
        foreach (var delay in delays)
        {
            await Task.Delay(delay, cancellationToken);
            using var req = new HttpRequestMessage(HttpMethod.Get, monitorUrl);
            using var resp = await client.SendAsync(req, cancellationToken);
            var body = await resp.Content.ReadAsStringAsync(cancellationToken);
            if (!resp.IsSuccessStatusCode && resp.StatusCode != System.Net.HttpStatusCode.Accepted)
            {
                return ("monitorFailed", $"{(int)resp.StatusCode}: {body}", null);
            }
            try
            {
                using var doc = JsonDocument.Parse(body);
                var status = doc.RootElement.TryGetProperty("status", out var s) ? s.GetString() : null;
                if (status is null) continue;
                if (string.Equals(status, "completed", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(status, "succeeded", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(status, "failed", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(status, "cancelled", StringComparison.OrdinalIgnoreCase))
                {
                    var error = doc.RootElement.TryGetProperty("error", out var e) ? e.ToString() : null;
                    var resource = doc.RootElement.TryGetProperty("resourceId", out var r) ? r.GetString() : null;
                    return (status, error, resource);
                }
            }
            catch
            {
                // Body wasn't JSON; keep polling.
            }
        }

        return ("timeout", "Graph copy did not complete within the polling window. Check the SharePoint destination shortly.", null);
    }
}

public enum WorkflowStatus
{
    Draft,
    InReview,
    Approved,
    Published,
}

