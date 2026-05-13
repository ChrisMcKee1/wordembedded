using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph.Models.ODataErrors;

namespace Wordembedded.Api.Middleware;

public sealed class GraphErrorMappingMiddleware(RequestDelegate next, ILogger<GraphErrorMappingMiddleware> logger)
{
    public async Task InvokeAsync(HttpContext context)
    {
        try
        {
            await next(context);
        }
        catch (ODataError ex) when (!context.Response.HasStarted)
        {
            var statusCode = ex.ResponseStatusCode > 0
                ? ex.ResponseStatusCode
                : StatusCodes.Status502BadGateway;

            logger.LogWarning(
                ex,
                "Microsoft Graph request failed with {StatusCode}: {GraphErrorCode} — {GraphErrorMessage}",
                statusCode,
                ex.Error?.Code,
                ex.Error?.Message);

            context.Response.StatusCode = statusCode;
            context.Response.ContentType = "application/problem+json";

            var problem = new ProblemDetails
            {
                Status = statusCode,
                Title = ex.Error?.Code ?? "Microsoft Graph request failed",
                Detail = ex.Error?.Message ?? ex.Message,
                Type = "https://learn.microsoft.com/graph/errors"
            };

            await context.Response.WriteAsJsonAsync(problem);
        }
    }
}
