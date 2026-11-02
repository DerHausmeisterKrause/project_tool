namespace TaskTool.Services;

public static class TicketCloseReplyComposer
{
    public const string CloseNotice = "Ist erfolgt. Ticket wird geschlossen.";

    public static string EnsureCloseNotice(string? renderedTemplate)
    {
        var normalized = (renderedTemplate ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Trim();
        if (normalized.Contains(CloseNotice, StringComparison.Ordinal))
            return normalized;
        if (normalized.Length == 0)
            return CloseNotice;

        var lines = normalized.Split('\n');
        var firstContent = Array.FindIndex(lines, line => !string.IsNullOrWhiteSpace(line));
        if (firstContent < 0)
            return CloseNotice;

        var before = lines.Take(firstContent + 1).Select(line => line.TrimEnd());
        var after = lines.Skip(firstContent + 1).SkipWhile(string.IsNullOrWhiteSpace).Select(line => line.TrimEnd());
        return string.Join(Environment.NewLine, before
            .Append(string.Empty)
            .Append(CloseNotice)
            .Append(string.Empty)
            .Concat(after)).TrimEnd();
    }
}
