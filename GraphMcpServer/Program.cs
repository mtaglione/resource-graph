using Azure.Identity;
using GraphMcpServer;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using ModelContextProtocol.Server;
using System.Security.Claims;

var builder = WebApplication.CreateBuilder(args);

// Route all logging to stderr so stdout stays clean for the MCP stdio protocol.
builder.Logging.ClearProviders();
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);
builder.Logging.SetMinimumLevel(LogLevel.Warning);

// Pull settings (including Key Vault references) from Azure App Configuration.
// In App Service the connection string arrives via a Key Vault reference app setting.
// Locally, DefaultAzureCredential (Azure CLI) is used against the endpoint directly.
var credential = new DefaultAzureCredential();

var appConfigConnectionString = builder.Configuration["AppConfiguration:ConnectionString"];
var appConfigEndpoint = builder.Configuration["AppConfiguration:Endpoint"];

if (string.IsNullOrWhiteSpace(appConfigConnectionString) && string.IsNullOrWhiteSpace(appConfigEndpoint))
    throw new InvalidOperationException("Either AppConfiguration:ConnectionString or AppConfiguration:Endpoint must be set.");

builder.Configuration.AddAzureAppConfiguration(options =>
{
    if (!string.IsNullOrWhiteSpace(appConfigConnectionString))
        options.Connect(appConfigConnectionString);
    else
        options.Connect(new Uri(appConfigEndpoint!), credential);

    options.ConfigureKeyVault(kv => kv.SetCredential(credential));
});

// JWT Bearer: accept tokens issued for this app but only from the jbb.com tenant.
// The app registration lives in jbb-devops.com (multi-tenant); callers from any
// tenant can obtain a token, but we reject anyone whose 'tid' claim is not jbb.com.
const string jbbTenantId = "eb410fdb-3a38-4552-909a-b515a547471c";
var mcpClientId = builder.Configuration["McpAuth:ClientId"]
    ?? throw new InvalidOperationException("McpAuth:ClientId must be configured.");

builder.Services
    .AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        // Use the /common endpoint so tokens from any tenant are accepted at the
        // JWKS / signature-validation level; tenant restriction is enforced below.
        options.Authority = "https://login.microsoftonline.com/common/v2.0";
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = false,          // issuer varies per tenant — see OnTokenValidated
            ValidateAudience = true,
            ValidAudiences = [$"api://{mcpClientId}", mcpClientId],
        };
        options.Events = new JwtBearerEvents
        {
            OnTokenValidated = ctx =>
            {
                var tid = ctx.Principal?.FindFirstValue("tid");
                if (tid != jbbTenantId)
                    ctx.Fail($"Tenant '{tid}' is not permitted. Only jbb.com users may access this service.");
                return Task.CompletedTask;
            }
        };
    });

builder.Services.AddAuthorization();

// Register GraphHttpClient as a typed HttpClient (transient by default).
builder.Services.AddHttpClient<GraphHttpClient>();

// Register the MCP server and discover tools in this assembly.
var mcpBuilder = builder.Services
    .AddMcpServer()
    .WithHttpTransport()
    .WithToolsFromAssembly();

// Azure App Service requires an ASP.NET Core web host listening on HTTP.
// Keep stdio transport enabled for local MCP usage outside App Service.
var isAppService = !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("WEBSITE_SITE_NAME"));
if (!isAppService)
{
    mcpBuilder.WithStdioServerTransport();
}

var app = builder.Build();

app.UseStaticFiles();
app.UseAuthentication();
app.UseAuthorization();

// MCP endpoints require a valid jbb.com JWT.
app.MapMcp().RequireAuthorization();

app.MapGet("/visualize", (IWebHostEnvironment env) =>
    Results.File(Path.Combine(env.WebRootPath, "visualize.html"), "text/html"))
    .AllowAnonymous();

app.MapGet("/info", () => Results.Ok(new { service = "GraphMcpServer", status = "running" }));
app.MapGet("/health", () => Results.Ok(new { status = "healthy" }));

await app.RunAsync();
