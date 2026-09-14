namespace TaskTool.Models;

public sealed class ZnunyAgent
{
    public int UserId { get; init; }
    public string Login { get; init; } = string.Empty;
    public string FirstName { get; init; } = string.Empty;
    public string LastName { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;

    public string DisplayName
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(Name)) return Name.Trim();
            var fullName = $"{FirstName} {LastName}".Trim();
            if (!string.IsNullOrWhiteSpace(fullName)) return fullName;
            return !string.IsNullOrWhiteSpace(Login) ? Login.Trim() : $"Agent {UserId}";
        }
    }

    public override string ToString() => DisplayName;
}

public sealed record TicketAssignmentUpdateResult(bool Success, string Message);
