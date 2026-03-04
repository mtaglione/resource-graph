using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

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
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    private readonly HttpClient _http;
    private readonly TokenCredential _credential;
    private readonly string _tenantId;
    private readonly string _appId;
    private readonly string _authMode;

    public GraphHttpClient(HttpClient http, IConfiguration config)
    {
        _http = http;
        _tenantId = FirstNonEmpty(
                        config["AzureAd:TenantId"],
                        config["TenantId"])
                    ?? throw new InvalidOperationException("TenantId is missing. Set AzureAd:TenantId or TenantId.");

        _appId = FirstNonEmpty(
                     config["ClientId"],
                     config["AzureAd:ClientId"])
                 ?? string.Empty;

        var clientSecret = FirstNonEmpty(
            config["ClientSecret"],
            config["AzureAd:ClientSecret"]);

        if (!string.IsNullOrWhiteSpace(clientSecret) && !string.IsNullOrWhiteSpace(_appId))
        {
            _credential = new ClientSecretCredential(
                tenantId: _tenantId,
                clientId: _appId,
                clientSecret: clientSecret);
            _authMode = "ClientSecretCredential";
        }
        else
        {
            var managedIdentityClientId = FirstNonEmpty(
                config["ManagedIdentityClientId"],
                config["AzureAd:ManagedIdentityClientId"]);

            var options = new DefaultAzureCredentialOptions
            {
                TenantId = _tenantId,
                ManagedIdentityClientId = managedIdentityClientId
            };

            _credential = new DefaultAzureCredential(options);
            _authMode = "DefaultAzureCredential";

            if (string.IsNullOrWhiteSpace(_appId) && !string.IsNullOrWhiteSpace(managedIdentityClientId))
            {
                _appId = managedIdentityClientId;
            }
        }
    }

    public string TenantId => _tenantId;
    public string AppId => _appId;
    public string AuthMode => _authMode;

    public Task<JsonObject> GetAsync(string endpoint, CancellationToken ct = default) =>
        SendJsonAsync(HttpMethod.Get, endpoint, contentFactory: null, ct);

    public Task<JsonObject> DeleteAsync(string endpoint, CancellationToken ct = default) =>
        SendJsonAsync(HttpMethod.Delete, endpoint, contentFactory: null, ct);

    public Task<JsonObject> PostAsync(string endpoint, JsonNode jsonBody, CancellationToken ct = default) =>
        SendJsonAsync(HttpMethod.Post, endpoint, () => JsonContent(jsonBody), ct);

    public Task<JsonObject> PatchAsync(string endpoint, JsonNode jsonBody, CancellationToken ct = default) =>
        SendJsonAsync(HttpMethod.Patch, endpoint, () => JsonContent(jsonBody), ct);

    public Task<JsonObject> PutBytesAsync(string endpoint, byte[] content, string mediaType = "application/octet-stream", CancellationToken ct = default) =>
        SendJsonAsync(HttpMethod.Put, endpoint, () =>
        {
            var byteContent = new ByteArrayContent(content);
            byteContent.Headers.ContentType = new MediaTypeHeaderValue(mediaType);
            return byteContent;
        }, ct);

    private async Task<JsonObject> SendJsonAsync(
        HttpMethod method,
        string endpoint,
        Func<HttpContent>? contentFactory,
        CancellationToken ct)
    {
        var attempt = 0;
        var delayMs = 500;

        while (true)
        {
            ct.ThrowIfCancellationRequested();
            using var request = new HttpRequestMessage(method, Resolve(endpoint));
            request.Headers.Authorization = await GetAuthHeaderAsync(ct);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            if (contentFactory is not null)
            {
                request.Content = contentFactory();
            }

            using var response = await _http.SendAsync(request, ct);

            if (response.IsSuccessStatusCode)
            {
                if (response.Content.Headers.ContentLength == 0)
                {
                    return new JsonObject { ["status"] = "ok" };
                }

                var raw = await response.Content.ReadAsStringAsync(ct);
                if (string.IsNullOrWhiteSpace(raw))
                {
                    return new JsonObject { ["status"] = "ok" };
                }

                var parsed = JsonNode.Parse(raw)?.AsObject();
                return parsed ?? new JsonObject { ["raw"] = raw };
            }

            if ((response.StatusCode == HttpStatusCode.TooManyRequests || response.StatusCode == HttpStatusCode.ServiceUnavailable)
                && attempt < 5)
            {
                var retryMs = RetryDelayMs(response, delayMs);
                await Task.Delay(retryMs, ct);
                attempt++;
                delayMs = Math.Min(delayMs * 2, 8000);
                continue;
            }

            throw await GraphApiException.FromResponseAsync(response, ct);
        }
    }

    private static int RetryDelayMs(HttpResponseMessage response, int currentDelayMs)
    {
        if (response.Headers.RetryAfter?.Delta is { } delta)
        {
            return (int)Math.Clamp(delta.TotalMilliseconds, 250, 30_000);
        }

        return currentDelayMs;
    }

    private async Task<AuthenticationHeaderValue> GetAuthHeaderAsync(CancellationToken ct)
    {
        var token = await _credential.GetTokenAsync(new TokenRequestContext(GraphScopes), ct);
        return new AuthenticationHeaderValue("Bearer", token.Token);
    }

    private static string Resolve(string endpoint) =>
        endpoint.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            ? endpoint
            : $"{GraphBase}/{endpoint.TrimStart('/')}";

    private static StringContent JsonContent(JsonNode body) =>
        new(body.ToJsonString(JsonOptions), Encoding.UTF8, "application/json");

    private static string? FirstNonEmpty(params string?[] values) =>
        values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
}

public sealed class GraphApiException : Exception
{
    public int Status { get; }
    public string Code { get; }
    public string? RequestId { get; }
    public JsonObject? InnerError { get; }

    public GraphApiException(int status, string code, string message, string? requestId, JsonObject? innerError)
        : base(message)
    {
        Status = status;
        Code = code;
        RequestId = requestId;
        InnerError = innerError;
    }

    public static async Task<GraphApiException> FromResponseAsync(HttpResponseMessage response, CancellationToken ct)
    {
        var status = (int)response.StatusCode;
        var requestId = TryHeader(response, "request-id") ?? TryHeader(response, "x-ms-request-id");
        var fallbackMessage = $"Microsoft Graph request failed with status {status}.";

        if (response.Content is null)
        {
            return new GraphApiException(status, "GraphRequestFailed", fallbackMessage, requestId, null);
        }

        var raw = await response.Content.ReadAsStringAsync(ct);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return new GraphApiException(status, "GraphRequestFailed", fallbackMessage, requestId, null);
        }

        try
        {
            var root = JsonNode.Parse(raw)?.AsObject();
            var error = root?["error"]?.AsObject();

            var code = error?["code"]?.GetValue<string>() ?? "GraphRequestFailed";
            var message = error?["message"]?.GetValue<string>() ?? fallbackMessage;
            var innerError = error?["innerError"] as JsonObject;
            var req = requestId
                      ?? innerError?["request-id"]?.GetValue<string>()
                      ?? innerError?["requestId"]?.GetValue<string>();

            return new GraphApiException(status, code, message, req, innerError);
        }
        catch
        {
            return new GraphApiException(status, "GraphRequestFailed", fallbackMessage, requestId, null);
        }
    }

    private static string? TryHeader(HttpResponseMessage response, string headerName)
    {
        if (response.Headers.TryGetValues(headerName, out var values))
        {
            return values.FirstOrDefault();
        }

        return null;
    }
}
