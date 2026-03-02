using ModelContextProtocol.Server;
using System.ComponentModel;

namespace GraphMcpServer.Tools;

[McpServerToolType]
public sealed class GraphTools(GraphHttpClient graph)
{
    [McpServerTool]
    [Description(
        "Execute a GET request against the Microsoft Graph API. " +
        "Accepts either a relative path (e.g. 'users', 'me', 'sites/root/lists') " +
        "or a full URL. Returns the raw JSON response.")]
    public Task<string> GraphGet(
        [Description("Graph API endpoint path or full URL (e.g. 'users', 'sites/{siteId}', 'https://graph.microsoft.com/v1.0/me').")]
        string endpoint,
        CancellationToken cancellationToken) =>
        graph.GetAsync(endpoint, cancellationToken);

    [McpServerTool]
    [Description(
        "Execute a POST request against the Microsoft Graph API. " +
        "Accepts either a relative path or full URL, and a JSON body. " +
        "Returns the raw JSON response.")]
    public Task<string> GraphPost(
        [Description("Graph API endpoint path or full URL.")]
        string endpoint,
        [Description("JSON body to include in the request.")]
        string body,
        CancellationToken cancellationToken) =>
        graph.PostAsync(endpoint, body, cancellationToken);
}
