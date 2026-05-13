namespace Wordembedded.Api.Options;

public sealed class SharePointEmbeddedOptions
{
    public required string ContainerTypeId { get; set; }
}

public sealed class WebhookOptions
{
    public required string NotificationUrl { get; set; }
    public required string ClientStateSecret { get; set; }
}
