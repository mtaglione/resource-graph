using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using System.Net.Http.Headers;

namespace GraphMcpServer;

/// <summary>
/// Thin HTTP wrapper around the Microsoft Graph REST API.
/// Authenticates using an app registration client credential and
/// resolves relative paths against https://graph.microsoft.com/v1.0/.
/// </summary>
public sealed class GraphHttpClient
{
    private const string GraphBase = "https://graph.microsoft.com/v1.0";
    private static readonly string[] GraphScopes = ["https://graph.microsoft.com/.default"];

    private readonly HttpClient _http;
    private readonly TokenCredential _credential;

    public GraphHttpClient(HttpClient http, IConfiguration config)
    {
        _http = http;
        _credential = new ClientSecretCredential(
            tenantId:     config["AzureAd:TenantId"]
                          ?? throw new InvalidOperationException("AzureAd:TenantId is missing"),
            clientId:     config["ClientId"]
                          ?? throw new InvalidOperationException("ClientId is missing"),
            clientSecret: config["ClientSecret"]
                          ?? throw new InvalidOperationException("ClientSecret is missing"));
    }

    public async Task<string> GetAsync(string endpoint, CancellationToken ct = default)
    {
        await SetAuthHeaderAsync(ct);
        var response = await _http.GetAsync(Resolve(endpoint), ct);
        return await response.Content.ReadAsStringAsync(ct);
    }

    public async Task<string> PostAsync(string endpoint, string jsonBody, CancellationToken ct = default)
    {
        await SetAuthHeaderAsync(ct);
        using var content = new StringContent(jsonBody, System.Text.Encoding.UTF8, "application/json");
        var response = await _http.PostAsync(Resolve(endpoint), content, ct);
        return await response.Content.ReadAsStringAsync(ct);
    }

    // -----------------------------------------------------------------------

    private async Task SetAuthHeaderAsync(CancellationToken ct)
    {
        var token = await _credential.GetTokenAsync(new TokenRequestContext(GraphScopes), ct);
        _http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", token.Token);
    }

    private static string Resolve(string endpoint) =>
        endpoint.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            ? endpoint
            : $"{GraphBase}/{endpoint.TrimStart('/')}";
}
