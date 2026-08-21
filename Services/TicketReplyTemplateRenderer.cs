using System.Net.Mail;
using System.Text.RegularExpressions;
using TaskTool.Models;

namespace TaskTool.Services;

public static class TicketReplyTemplateRenderer
{
    private const int MaximumTemplateLength = 10000;
    private static readonly Regex VariablePattern = new(
        @"\$(CustomerFirstName|CustomerLastName|TicketNumber|TicketTitle|Customer|Date)(?![A-Za-z0-9_])",
        RegexOptions.CultureInvariant);

    public static TicketReplyTemplateContext CreateContext(
        TicketArticleItem? originalCustomerArticle,
        string ticketNumber,
        string ticketTitle,
        DateTime localDate)
    {
        var customer = GetCustomerDisplayName(originalCustomerArticle?.From);
        var (firstName, lastName) = SplitCustomerName(customer);
        return new TicketReplyTemplateContext(
            customer,
            firstName,
            lastName,
            ticketNumber ?? string.Empty,
            ticketTitle ?? string.Empty,
            localDate);
    }

    public static TicketReplyTemplateResult Render(string? template, TicketReplyTemplateContext context)
    {
        var source = template ?? string.Empty;
        if (source.Length > MaximumTemplateLength)
            source = source[..MaximumTemplateLength];

        var values = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["$Customer"] = context.Customer,
            ["$CustomerFirstName"] = context.CustomerFirstName,
            ["$CustomerLastName"] = context.CustomerLastName,
            ["$TicketNumber"] = context.TicketNumber,
            ["$TicketTitle"] = context.TicketTitle,
            ["$Date"] = context.LocalDate.ToString("dd.MM.yyyy")
        };
        var rendered = VariablePattern.Replace(source, match =>
            values.TryGetValue(match.Value, out var value) && !string.IsNullOrWhiteSpace(value)
                ? value
                : match.Value);
        if (rendered.Length > MaximumTemplateLength)
            rendered = rendered[..MaximumTemplateLength];
        return new TicketReplyTemplateResult(rendered, VariablePattern.IsMatch(rendered));
    }

    private static string GetCustomerDisplayName(string? from)
    {
        if (string.IsNullOrWhiteSpace(from)) return string.Empty;
        try
        {
            return new MailAddress(from.Trim()).DisplayName.Trim();
        }
        catch (FormatException)
        {
            return string.Empty;
        }
    }

    private static (string firstName, string lastName) SplitCustomerName(string customer)
    {
        if (string.IsNullOrWhiteSpace(customer)) return (string.Empty, string.Empty);
        if (customer.Count(character => character == ',') == 1)
        {
            var parts = customer.Split(',', 2, StringSplitOptions.TrimEntries);
            if (parts.Length == 2 && parts.All(part => !string.IsNullOrWhiteSpace(part)))
                return (parts[1].Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty, parts[0]);
        }

        var words = customer.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return words.Length >= 2 ? (words[0], words[^1]) : (string.Empty, string.Empty);
    }
}
