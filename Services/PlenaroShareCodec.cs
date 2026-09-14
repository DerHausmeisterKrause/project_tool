using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using TaskTool.Models;

namespace TaskTool.Services;

public static class PlenaroShareCodec
{
    public const string BeginMarker = "--- PLENARO-SHARE BEGIN ---";
    public const string EndMarker = "--- PLENARO-SHARE END ---";
    private static readonly Regex BlockRegex = new(Regex.Escape(BeginMarker) + ".*?" + Regex.Escape(EndMarker), RegexOptions.Singleline);

    public static string AppendOrReplace(string humanBody, PlenaroSharePayloadV1 payload, out string hash)
    {
        var json = JsonSerializer.SerializeToUtf8Bytes(payload);
        hash = Convert.ToHexString(SHA256.HashData(json));
        var encoded = Convert.ToBase64String(json).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        var clean = BlockRegex.Replace(humanBody ?? string.Empty, string.Empty).TrimEnd();
        return $"{clean}\n\n{BeginMarker}\nVersion: 1\nPayload: {encoded}\nSHA256: {hash}\n{EndMarker}";
    }

    public static bool TryParse(string body, out PlenaroSharePayloadV1? payload, out string hash, out string reason)
    {
        payload = null; hash = string.Empty; reason = "invalid-block";
        var matches = BlockRegex.Matches(body ?? string.Empty);
        if (matches.Count != 1) { reason = matches.Count > 1 ? "ambiguous-blocks" : "missing-block"; return false; }
        var lines = matches[0].Value.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var version = Value(lines, "Version:");
        var encoded = Value(lines, "Payload:");
        hash = Value(lines, "SHA256:");
        if (version != "1") { reason = "unknown-version"; return false; }
        try
        {
            var padded = encoded.Replace('-', '+').Replace('_', '/');
            padded = padded.PadRight((padded.Length + 3) / 4 * 4, '=');
            var bytes = Convert.FromBase64String(padded);
            var actual = Convert.ToHexString(SHA256.HashData(bytes));
            var expectedBytes = Encoding.ASCII.GetBytes(hash.ToUpperInvariant());
            var actualBytes = Encoding.ASCII.GetBytes(actual);
            if (expectedBytes.Length != actualBytes.Length || !CryptographicOperations.FixedTimeEquals(actualBytes, expectedBytes)) { reason = "checksum-invalid"; return false; }
            payload = JsonSerializer.Deserialize<PlenaroSharePayloadV1>(bytes);
            if (payload == null || payload.Version != 1) { reason = "unknown-version"; return false; }
            if (!Guid.TryParse(payload.TaskShareId, out _) || !Guid.TryParse(payload.SegmentShareId, out _)) { reason = "missing-share-id"; return false; }
            if (payload.SegmentStartLocal == default || payload.SegmentEndLocal <= payload.SegmentStartLocal || payload.SegmentEndLocal - payload.SegmentStartLocal > TimeSpan.FromDays(7)) { reason = "invalid-time-range"; return false; }
            reason = string.Empty; return true;
        }
        catch (FormatException) { reason = "base64-invalid"; return false; }
        catch (JsonException) { reason = "json-invalid"; return false; }
    }

    private static string Value(string[] lines, string prefix) => lines.FirstOrDefault(x => x.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))?[prefix.Length..].Trim() ?? string.Empty;
}
