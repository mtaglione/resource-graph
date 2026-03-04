using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text.Json.Nodes;

namespace GraphMcpServer.Tools;

[McpServerToolType]
public sealed class GraphTools(GraphHttpClient graph)
{
    private const int DefaultTop = 50;
    private const int MaxTop = 50;

    [McpServerTool]
    [Description("Search sites by keyword using Microsoft Graph /sites?search=...")]
    public Task<JsonObject> sp_sites_search(
        [Description("Search query text.")] string query,
        [Description("Max rows to return (default 50, max 50). Use cursor for next page.")] int? top,
        [Description("Opaque paging cursor returned from prior call.")] string? cursor,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var endpoint = BuildPagedEndpoint(
                "sites",
                top,
                cursor,
                new Dictionary<string, string?>
                {
                    ["search"] = query
                });

            return await QueryCollectionAsync(endpoint, ct);
        }, cancellationToken);

    [McpServerTool]
    [Description("Get a site by siteId or by hostname + sitePath.")]
    public Task<JsonObject> sp_sites_get(
        [Description("Site ID. Optional if hostname + sitePath are provided.")] string? siteId,
        [Description("Host name like contoso.sharepoint.com.")] string? hostname,
        [Description("Server-relative site path like /sites/Engineering.")] string? sitePath,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            string endpoint;

            if (!string.IsNullOrWhiteSpace(siteId))
            {
                endpoint = $"sites/{siteId}";
            }
            else if (!string.IsNullOrWhiteSpace(hostname) && !string.IsNullOrWhiteSpace(sitePath))
            {
                var normalizedPath = sitePath.StartsWith('/') ? sitePath : $"/{sitePath}";
                endpoint = $"sites/{hostname}:{normalizedPath}";
            }
            else
            {
                throw new InvalidOperationException("Provide siteId OR hostname + sitePath.");
            }

            var result = await graph.GetAsync(endpoint, ct);
            return new JsonObject { ["data"] = result };
        }, cancellationToken);

    [McpServerTool]
    [Description("List lists for a site using /sites/{site-id}/lists.")]
    public Task<JsonObject> sp_lists_list(
        [Description("SharePoint site ID.")] string siteId,
        [Description("Max rows to return (default 50, max 50). Use cursor for next page.")] int? top,
        [Description("Opaque paging cursor returned from prior call.")] string? cursor,
        CancellationToken cancellationToken) =>
        RunTool(ct => QueryCollectionAsync(
            BuildPagedEndpoint($"sites/{siteId}/lists", top, cursor),
            ct), cancellationToken);

    [McpServerTool]
    [Description("Get list metadata using /sites/{site-id}/lists/{list-id}.")]
    public Task<JsonObject> sp_list_get(
        [Description("SharePoint site ID.")] string siteId,
        [Description("List ID.")] string listId,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var data = await graph.GetAsync($"sites/{siteId}/lists/{listId}", ct);
            return new JsonObject { ["data"] = data };
        }, cancellationToken);

    [McpServerTool]
    [Description("List columns for a list using /sites/{site-id}/lists/{list-id}/columns.")]
    public Task<JsonObject> sp_list_columns_list(
        [Description("SharePoint site ID.")] string siteId,
        [Description("List ID.")] string listId,
        [Description("Max rows to return (default 50, max 50). Use cursor for next page.")] int? top,
        [Description("Opaque paging cursor returned from prior call.")] string? cursor,
        CancellationToken cancellationToken) =>
        RunTool(ct => QueryCollectionAsync(
            BuildPagedEndpoint($"sites/{siteId}/lists/{listId}/columns", top, cursor),
            ct), cancellationToken);

    [McpServerTool]
    [Description("List items in a list, optionally expanding fields.")]
    public Task<JsonObject> sp_list_items_list(
        [Description("SharePoint site ID.")] string siteId,
        [Description("List ID.")] string listId,
        [Description("Whether to expand fields on each list item.")] bool expandFields,
        [Description("OData filter expression.")] string? filter,
        [Description("OData orderby expression.")] string? orderby,
        [Description("OData select fields for item entity.")] string? select,
        [Description("Max rows to return (default 50, max 50). Use cursor for next page.")] int? top,
        [Description("Opaque paging cursor returned from prior call.")] string? cursor,
        CancellationToken cancellationToken) =>
        RunTool(ct =>
        {
            var query = new Dictionary<string, string?>
            {
                ["$filter"] = filter,
                ["$orderby"] = orderby,
                ["$select"] = select,
                ["$expand"] = expandFields ? "fields" : null
            };

            return QueryCollectionAsync(
                BuildPagedEndpoint($"sites/{siteId}/lists/{listId}/items", top, cursor, query),
                ct);
        }, cancellationToken);

    [McpServerTool]
    [Description("Get a single list item and expand fields.")]
    public Task<JsonObject> sp_list_item_get(
        [Description("SharePoint site ID.")] string siteId,
        [Description("List ID.")] string listId,
        [Description("List item ID.")] string itemId,
        [Description("OData select fields for item entity.")] string? select,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var endpoint = BuildEndpoint(
                $"sites/{siteId}/lists/{listId}/items/{itemId}",
                new Dictionary<string, string?>
                {
                    ["$expand"] = "fields",
                    ["$select"] = select
                });

            var data = await graph.GetAsync(endpoint, ct);
            return new JsonObject { ["data"] = data };
        }, cancellationToken);

    [McpServerTool]
    [Description("Create a list item with fields.")]
    public Task<JsonObject> sp_list_item_create(
        [Description("SharePoint site ID.")] string siteId,
        [Description("List ID.")] string listId,
        [Description("Field values object as JSON text.")] string fields,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var fieldsObject = ParseJsonObject(fields, "fields");
            var body = new JsonObject { ["fields"] = fieldsObject };
            var data = await graph.PostAsync($"sites/{siteId}/lists/{listId}/items", body, ct);
            return new JsonObject { ["data"] = data };
        }, cancellationToken);

    //[McpServerTool]
    //[Description("Update list item fields.")]
    //public Task<JsonObject> sp_list_item_update(
    //    [Description("SharePoint site ID.")] string siteId,
    //    [Description("List ID.")] string listId,
    //    [Description("List item ID.")] string itemId,
    //    [Description("Field values object.")] JsonObject fields,
    //    CancellationToken cancellationToken) =>
    //    RunTool(async ct =>
    //    {
    //        var data = await graph.PatchAsync($"sites/{siteId}/lists/{listId}/items/{itemId}/fields", fields, ct);
    //        return new JsonObject { ["data"] = data };
    //    }, cancellationToken);

    // [McpServerTool]
    // [Description("Delete a list item.")]
    // public Task<JsonObject> sp_list_item_delete(
    //     [Description("SharePoint site ID.")] string siteId,
    //     [Description("List ID.")] string listId,
    //     [Description("List item ID.")] string itemId,
    //     CancellationToken cancellationToken) =>
    //     RunTool(async ct =>
    //     {
    //         var data = await graph.DeleteAsync($"sites/{siteId}/lists/{listId}/items/{itemId}", ct);
    //         return new JsonObject
    //         {
    //             ["data"] = data,
    //             ["deleted"] = true
    //         };
    //     }, cancellationToken);

    [McpServerTool]
    [Description("List drives for a site.")]
    public Task<JsonObject> sp_drives_list(
        [Description("SharePoint site ID.")] string siteId,
        [Description("Max rows to return (default 50, max 50). Use cursor for next page.")] int? top,
        [Description("Opaque paging cursor returned from prior call.")] string? cursor,
        CancellationToken cancellationToken) =>
        RunTool(ct => QueryCollectionAsync(
            BuildPagedEndpoint($"sites/{siteId}/drives", top, cursor),
            ct), cancellationToken);

    [McpServerTool]
    [Description("Get drive metadata.")]
    public Task<JsonObject> sp_drive_get(
        [Description("Drive ID.")] string driveId,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var data = await graph.GetAsync($"drives/{driveId}", ct);
            return new JsonObject { ["data"] = data };
        }, cancellationToken);

    [McpServerTool]
    [Description("List children of a drive root.")]
    public Task<JsonObject> sp_drive_root_list_children(
        [Description("Drive ID.")] string driveId,
        [Description("Max rows to return (default 50, max 50). Use cursor for next page.")] int? top,
        [Description("Opaque paging cursor returned from prior call.")] string? cursor,
        CancellationToken cancellationToken) =>
        RunTool(ct => QueryCollectionAsync(
            BuildPagedEndpoint($"drives/{driveId}/root/children", top, cursor),
            ct), cancellationToken);

    [McpServerTool]
    [Description("List children of a folder drive item.")]
    public Task<JsonObject> sp_driveitem_list_children(
        [Description("Drive ID.")] string driveId,
        [Description("Parent folder item ID.")] string itemId,
        [Description("Max rows to return (default 50, max 50). Use cursor for next page.")] int? top,
        [Description("Opaque paging cursor returned from prior call.")] string? cursor,
        CancellationToken cancellationToken) =>
        RunTool(ct => QueryCollectionAsync(
            BuildPagedEndpoint($"drives/{driveId}/items/{itemId}/children", top, cursor),
            ct), cancellationToken);

    [McpServerTool]
    [Description("Get drive item metadata.")]
    public Task<JsonObject> sp_driveitem_get(
        [Description("Drive ID.")] string driveId,
        [Description("Item ID.")] string itemId,
        [Description("OData select expression.")] string? select,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var endpoint = BuildEndpoint(
                $"drives/{driveId}/items/{itemId}",
                new Dictionary<string, string?>
                {
                    ["$select"] = select
                });

            var data = await graph.GetAsync(endpoint, ct);
            return new JsonObject { ["data"] = data };
        }, cancellationToken);

    [McpServerTool]
    [Description("Get short-lived download URL for a drive item (metadata only, no file bytes).")]
    public Task<JsonObject> sp_driveitem_download_url(
        [Description("Drive ID.")] string driveId,
        [Description("Item ID.")] string itemId,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var endpoint = BuildEndpoint(
                $"drives/{driveId}/items/{itemId}",
                new Dictionary<string, string?>
                {
                    ["$select"] = "id,name,size,file,folder,@microsoft.graph.downloadUrl"
                });

            var data = await graph.GetAsync(endpoint, ct);
            return new JsonObject
            {
                ["data"] = new JsonObject
                {
                    ["id"] = data["id"]?.DeepClone(),
                    ["name"] = data["name"]?.DeepClone(),
                    ["size"] = data["size"]?.DeepClone(),
                    ["downloadUrl"] = data["@microsoft.graph.downloadUrl"]?.DeepClone()
                }
            };
        }, cancellationToken);

    [McpServerTool]
    [Description("Upload a small file (simple upload) to a drive path using base64 content.")]
    public Task<JsonObject> sp_driveitem_upload_small(
        [Description("Drive ID.")] string driveId,
        [Description("Destination path under root (for example folder/file.txt).")]
        string path,
        [Description("Base64 encoded file content.")] string contentBase64,
        [Description("Conflict behavior: fail, replace, or rename.")] string conflictBehavior,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var content = Convert.FromBase64String(contentBase64);
            var safePath = path.TrimStart('/');
            var endpoint = BuildEndpoint(
                $"drives/{driveId}/root:/{safePath}:/content",
                new Dictionary<string, string?>
                {
                    ["@microsoft.graph.conflictBehavior"] = NormalizeConflictBehavior(conflictBehavior)
                });

            var data = await graph.PutBytesAsync(endpoint, content, "application/octet-stream", ct);
            return new JsonObject { ["data"] = data };
        }, cancellationToken);

    [McpServerTool]
    [Description("Create a folder under a parent drive item.")]
    public Task<JsonObject> sp_driveitem_create_folder(
        [Description("Drive ID.")] string driveId,
        [Description("Parent item ID.")] string parentId,
        [Description("Folder name.")] string folderName,
        [Description("Conflict behavior: fail, replace, or rename.")] string conflictBehavior,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var body = new JsonObject
            {
                ["name"] = folderName,
                ["folder"] = new JsonObject(),
                ["@microsoft.graph.conflictBehavior"] = NormalizeConflictBehavior(conflictBehavior)
            };

            var data = await graph.PostAsync($"drives/{driveId}/items/{parentId}/children", body, ct);
            return new JsonObject { ["data"] = data };
        }, cancellationToken);

    [McpServerTool]
    [Description("Search across SharePoint/Drive entities using Graph search API.")]
    public Task<JsonObject> sp_search(
        [Description("Search query text.")] string query,
        [Description("Entity types (default: driveItem,listItem,site).")]
        string[]? entityTypes,
        [Description("Max rows to return (default 50, max 50). Use cursor for next page.")] int? top,
        [Description("Opaque paging cursor returned from prior call.")] string? cursor,
        CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            var pageSize = ClampTop(top);
            var from = ParseCursorOffset(cursor);
            var entities = (entityTypes is { Length: > 0 }
                ? entityTypes
                : ["driveItem", "listItem", "site"])
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var body = new JsonObject
            {
                ["requests"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["entityTypes"] = new JsonArray(entities.Select(e => (JsonNode)e).ToArray()),
                        ["query"] = new JsonObject { ["queryString"] = query },
                        ["from"] = from,
                        ["size"] = pageSize
                    }
                }
            };

            var response = await graph.PostAsync("search/query", body, ct);

            var total = response["value"]?[0]?["hitsContainers"]?[0]?["total"]?.GetValue<int?>() ?? 0;
            var hits = response["value"]?[0]?["hitsContainers"]?[0]?["hits"] as JsonArray ?? [];
            var nextFrom = from + hits.Count;
            var nextCursor = nextFrom < total ? nextFrom.ToString() : null;

            return new JsonObject
            {
                ["value"] = hits,
                ["top"] = pageSize,
                ["cursor"] = nextCursor,
                ["total"] = total
            };
        }, cancellationToken);

    [McpServerTool]
    [Description("Returns identity context used for Graph token acquisition.")]
    public Task<JsonObject> sp_whoami(CancellationToken cancellationToken) =>
        RunTool(async ct =>
        {
            string? userDisplayName = null;
            try
            {
                var me = await graph.GetAsync("me?$select=displayName", ct);
                userDisplayName = me["displayName"]?.GetValue<string>();
            }
            catch
            {
                userDisplayName = null;
            }

            return new JsonObject
            {
                ["data"] = new JsonObject
                {
                    ["tenant"] = graph.TenantId,
                    ["appId"] = graph.AppId,
                    ["authMode"] = graph.AuthMode,
                    ["userDisplayName"] = userDisplayName
                }
            };
        }, cancellationToken);

    private async Task<JsonObject> QueryCollectionAsync(string endpoint, CancellationToken ct)
    {
        var result = await graph.GetAsync(endpoint, ct);
        var items = (result["value"] as JsonArray ?? []).DeepClone().AsArray();
        var next = result["@odata.nextLink"]?.GetValue<string>();

        return new JsonObject
        {
            ["value"] = items,
            ["top"] = items.Count,
            ["cursor"] = next
        };
    }

    private static string BuildPagedEndpoint(
        string basePath,
        int? top,
        string? cursor,
        Dictionary<string, string?>? extraQuery = null)
    {
        if (!string.IsNullOrWhiteSpace(cursor))
        {
            return cursor;
        }

        var query = extraQuery is null
            ? new Dictionary<string, string?>()
            : new Dictionary<string, string?>(extraQuery);
        query["$top"] = ClampTop(top).ToString();
        return BuildEndpoint(basePath, query);
    }

    private static string BuildEndpoint(string basePath, Dictionary<string, string?>? query = null)
    {
        if (query is null || query.Count == 0)
        {
            return basePath;
        }

        var parts = query
            .Where(kv => !string.IsNullOrWhiteSpace(kv.Value))
            .Select(kv => $"{Uri.EscapeDataString(kv.Key)}={Uri.EscapeDataString(kv.Value!)}")
            .ToArray();

        if (parts.Length == 0)
        {
            return basePath;
        }

        return $"{basePath}?{string.Join("&", parts)}";
    }

    private static int ClampTop(int? top)
    {
        var value = top.GetValueOrDefault(DefaultTop);
        if (value <= 0)
        {
            value = DefaultTop;
        }

        return Math.Min(value, MaxTop);
    }

    private static string NormalizeConflictBehavior(string value)
    {
        var normalized = value.Trim().ToLowerInvariant();
        return normalized switch
        {
            "fail" => "fail",
            "replace" => "replace",
            "rename" => "rename",
            _ => throw new InvalidOperationException("conflictBehavior must be one of: fail, replace, rename.")
        };
    }

    private static int ParseCursorOffset(string? cursor)
    {
        if (string.IsNullOrWhiteSpace(cursor))
        {
            return 0;
        }

        return int.TryParse(cursor, out var offset) && offset >= 0
            ? offset
            : throw new InvalidOperationException("Invalid cursor. Expected a non-negative numeric offset.");
    }

    private static JsonObject ParseJsonObject(string json, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            throw new InvalidOperationException($"{parameterName} must be a non-empty JSON object string.");
        }

        try
        {
            return JsonNode.Parse(json)?.AsObject()
                ?? throw new InvalidOperationException($"{parameterName} must be a JSON object.");
        }
        catch (Exception)
        {
            throw new InvalidOperationException($"{parameterName} must be valid JSON object text.");
        }
    }

    private static async Task<JsonObject> RunTool(Func<CancellationToken, Task<JsonObject>> action, CancellationToken ct)
    {
        try
        {
            return await action(ct);
        }
        catch (GraphApiException ex)
        {
            return GraphError(ex);
        }
        catch (Exception ex)
        {
            return new JsonObject
            {
                ["error"] = new JsonObject
                {
                    ["code"] = "InternalError",
                    ["message"] = ex.Message,
                    ["status"] = 500
                }
            };
        }
    }

    private static JsonObject GraphError(GraphApiException ex)
    {
        var error = new JsonObject
        {
            ["code"] = ex.Code,
            ["message"] = ex.Message,
            ["status"] = ex.Status
        };

        if (!string.IsNullOrWhiteSpace(ex.RequestId))
        {
            error["requestId"] = ex.RequestId;
        }

        if (ex.InnerError is not null)
        {
            error["innerError"] = ex.InnerError.DeepClone();
        }

        return new JsonObject { ["error"] = error };
    }
}
