using Microsoft.AspNetCore.Mvc;

namespace Wordembedded.Api.Endpoints;

internal static class EndpointResults
{
    public static IResult BadRequestProblem(string title, string detail) => Results.BadRequest(new ProblemDetails
    {
        Status = StatusCodes.Status400BadRequest,
        Title = title,
        Detail = detail
    });

    public static IResult NotFoundProblem(string title, string detail) => Results.NotFound(new ProblemDetails
    {
        Status = StatusCodes.Status404NotFound,
        Title = title,
        Detail = detail
    });
}
