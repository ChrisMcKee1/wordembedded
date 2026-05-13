using Microsoft.Graph;
using Microsoft.Identity.Web;
using Microsoft.Kiota.Abstractions.Authentication;
using Wordembedded.Api.Endpoints;
using Wordembedded.Api.Middleware;
using Wordembedded.Api.Options;
using Wordembedded.Api.Services;

var builder = WebApplication.CreateBuilder(args);

builder.AddServiceDefaults();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddCors(options =>
{
    options.AddPolicy("Frontend", policy =>
    {
        policy.WithOrigins("http://localhost:3000")
            .AllowAnyHeader()
            .AllowAnyMethod()
            .AllowCredentials();
    });
});

builder.Services
    .AddMicrosoftIdentityWebApiAuthentication(builder.Configuration, "AzureAd")
    .EnableTokenAcquisitionToCallDownstreamApi()
    .AddMicrosoftGraph(builder.Configuration.GetSection("DownstreamApi"))
    .AddInMemoryTokenCaches();

builder.Services.AddAuthorization();

builder.Services.AddOptions<SharePointEmbeddedOptions>()
    .Bind(builder.Configuration.GetSection("SharePointEmbedded"))
    .Validate(options => !string.IsNullOrWhiteSpace(options.ContainerTypeId), "SharePointEmbedded:ContainerTypeId is required.")
    .ValidateOnStart();

builder.Services.AddOptions<WebhookOptions>()
    .Bind(builder.Configuration.GetSection("Webhooks"))
    .Validate(options => !string.IsNullOrWhiteSpace(options.NotificationUrl), "Webhooks:NotificationUrl is required.")
    .Validate(options => !string.IsNullOrWhiteSpace(options.ClientStateSecret), "Webhooks:ClientStateSecret is required.")
    .ValidateOnStart();

builder.Services.AddHttpClient();

builder.Services.AddSingleton<IAppOnlyGraphClient>(sp =>
{
    var scopeFactory = sp.GetRequiredService<IServiceScopeFactory>();
    var authProvider = new BaseBearerTokenAuthenticationProvider(new AppOnlyAccessTokenProvider(scopeFactory));
    return new AppOnlyGraphClient(new GraphServiceClient(authProvider));
});

builder.Services.AddSingleton<IGraphSubscriptionStore, GraphSubscriptionStore>();
builder.Services.AddHostedService<WebhookRenewalService>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseCors("Frontend");
app.UseMiddleware<GraphErrorMappingMiddleware>();
app.UseAuthentication();
app.UseAuthorization();

app.MapDefaultEndpoints();

app.MapGet("/api/health", () => Results.Ok(new { status = "ok", version = "0.1.0" }))
    .AllowAnonymous();

app.MapContainerEndpoints();
app.MapFileEndpoints();
app.MapPreviewEndpoints();
app.MapSharingEndpoints();
app.MapVersionEndpoints();
app.MapSearchEndpoints();
app.MapWebhookEndpoints();
app.MapRecycleBinEndpoints();
app.MapThumbnailEndpoints();
app.MapSharePointBrowseEndpoints();
app.MapWorkflowEndpoints();

app.Run();
