using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions.Authentication;

namespace Wordembedded.Api.Services;

public interface IAppOnlyGraphClient
{
    GraphServiceClient Client { get; }
}

public sealed class AppOnlyGraphClient(GraphServiceClient client) : IAppOnlyGraphClient
{
    public GraphServiceClient Client { get; } = client;
}

public sealed class AppOnlyAccessTokenProvider(IServiceScopeFactory scopeFactory) : IAccessTokenProvider
{
    private const string GraphScope = "https://graph.microsoft.com/.default";

    public AllowedHostsValidator AllowedHostsValidator { get; } = new(new[] { "graph.microsoft.com" });

    public async Task<string> GetAuthorizationTokenAsync(
        Uri uri,
        Dictionary<string, object>? additionalAuthenticationContext = null,
        CancellationToken cancellationToken = default)
    {
        using var scope = scopeFactory.CreateScope();
        var tokenAcquisition = scope.ServiceProvider.GetRequiredService<ITokenAcquisition>();
        return await tokenAcquisition.GetAccessTokenForAppAsync(GraphScope);
    }
}
