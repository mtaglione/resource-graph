using Azure.Identity;
using GraphMcpServer;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

var builder = Host.CreateApplicationBuilder(args);

// Route all logging to stderr so stdout stays clean for the MCP stdio protocol.
builder.Logging.ClearProviders();
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);
builder.Logging.SetMinimumLevel(LogLevel.Warning);

// Pull ClientId + ClientSecret (Key Vault reference) from Azure App Configuration.
// DefaultAzureCredential picks up your logged-in Azure CLI identity locally,
// and a Managed Identity when deployed to Azure.
var appConfigEndpoint = builder.Configuration["AppConfiguration:Endpoint"]
    ?? throw new InvalidOperationException(
        "AppConfiguration:Endpoint is required in appsettings.json");

var azureCredential = new DefaultAzureCredential();

builder.Configuration.AddAzureAppConfiguration(options =>
    options
        .Connect(new Uri(appConfigEndpoint), azureCredential)
        .ConfigureKeyVault(kv => kv.SetCredential(azureCredential)));

// Register GraphHttpClient as a typed HttpClient (transient by default).
builder.Services.AddHttpClient<GraphHttpClient>();

// Register the MCP server with stdio transport and discover tools in this assembly.
builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithToolsFromAssembly();

await builder.Build().RunAsync();
