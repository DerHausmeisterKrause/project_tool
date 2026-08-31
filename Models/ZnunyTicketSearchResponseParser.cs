using System.Net;
using System.Text.Json;

namespace TaskTool.Models;

public enum ZnunyTicketSearchResponseShape
{
    TicketIds,
    EmptyObject
}

public sealed record ZnunyTicketSearchParseResult(
    IReadOnlyList<string> TicketIds,
    ZnunyTicketSearchResponseShape ResponseShape);

public static class ZnunyTicketSearchResponseParser
{
    public static ZnunyTicketSearchParseResult ExtractTicketIdsStrict(string responseBody, string stage)
    {
        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(responseBody);
        }
        catch (JsonException)
        {
            throw ProtocolError(stage, "TicketSearch response is not valid JSON.", responseBody);
        }

        using (document)
        {
            var root = document.RootElement;
            ThrowIfApiError(root, stage, responseBody);

            if (TryGetPropertyCaseInsensitive(root, "TicketIDs", out var ids)
                || TryGetPropertyCaseInsensitive(root, "TicketID", out ids))
            {
                return new ZnunyTicketSearchParseResult(
                    ExtractTicketIdValues(ids).ToList(),
                    ZnunyTicketSearchResponseShape.TicketIds);
            }

            // Znuny's TicketSearch operation serializes a successful search with no
            // matches as Data => {}, which the REST transport exposes as HTTP 200 + {}.
            if (root.ValueKind == JsonValueKind.Object && !root.EnumerateObject().Any())
            {
                return new ZnunyTicketSearchParseResult(
                    Array.Empty<string>(),
                    ZnunyTicketSearchResponseShape.EmptyObject);
            }
        }

        throw ProtocolError(stage,
            "TicketSearch response contains neither TicketID nor TicketIDs.", responseBody);
    }

    private static void ThrowIfApiError(JsonElement root, string stage, string responseBody)
    {
        if (!TryGetPropertyCaseInsensitive(root, "Error", out var error)
            || error.ValueKind != JsonValueKind.Object)
            return;

        throw new ZnunyApiException(
            stage,
            HttpStatusCode.OK,
            GetString(error, "ErrorCode"),
            GetString(error, "ErrorMessage"),
            responseBody);
    }

    private static IEnumerable<string> ExtractTicketIdValues(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in value.EnumerateArray())
            {
                var id = TicketIdToString(item);
                if (!string.IsNullOrWhiteSpace(id))
                    yield return id;
            }
            yield break;
        }

        var single = TicketIdToString(value);
        if (!string.IsNullOrWhiteSpace(single))
            yield return single;
    }

    private static string TicketIdToString(JsonElement value)
        => value.ValueKind is JsonValueKind.String or JsonValueKind.Number ? value.ToString() : string.Empty;

    private static string GetString(JsonElement root, string name)
        => TryGetPropertyCaseInsensitive(root, name, out var value) ? value.ToString() : string.Empty;

    private static bool TryGetPropertyCaseInsensitive(JsonElement root, string name, out JsonElement value)
    {
        if (root.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in root.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }

    private static ZnunyApiException ProtocolError(string stage, string message, string responseBody)
        => new(stage, HttpStatusCode.OK, "Protocol", message, responseBody);
}
