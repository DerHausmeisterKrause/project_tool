using System.Net.Mail;

namespace TaskTool.Services;

public static class SegmentAttendeeParser
{
    public static bool TryParse(string? text, out IReadOnlyList<string> addresses, out string invalidAddress)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var value in (text ?? string.Empty).Split(new[] { ';', ',', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            try
            {
                var parsed = new MailAddress(value);
                if (!string.Equals(parsed.Address, value, StringComparison.OrdinalIgnoreCase)) throw new FormatException();
                result.Add(parsed.Address);
            }
            catch (FormatException) { addresses = Array.Empty<string>(); invalidAddress = value; return false; }
        }
        addresses = result.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray(); invalidAddress = string.Empty; return true;
    }
}
