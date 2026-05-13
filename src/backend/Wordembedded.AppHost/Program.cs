#pragma warning disable ASPIREJAVASCRIPT001

var builder = DistributedApplication.CreateBuilder(args);

var api = builder.AddProject<Projects.Wordembedded_Api>("api")
    .WithHttpsEndpoint(port: 7071, name: "https")
    .WithHttpEndpoint(port: 5071, name: "http")
    .WithExternalHttpEndpoints();

builder.AddNextJsApp("frontend", "../../frontend", runScriptName: "dev")
    .WithHttpEndpoint(env: "PORT", port: 3000, name: "http")
    .WithEnvironment("NEXT_PUBLIC_API_BASE_URL", api.GetEndpoint("https"))
    .WithExternalHttpEndpoints();

builder.Build().Run();

