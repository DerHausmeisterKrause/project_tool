namespace TaskTool.Models;

public class TaskItem
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string TicketUrl { get; set; } = string.Empty;
    public DateTime? StartLocal { get; set; }
    public DateTime? EndLocal { get; set; }
    public TaskStatus Status { get; set; } = TaskStatus.Planned;
    public int? Priority { get; set; }
    public string Tags { get; set; } = string.Empty;
    public string OutlookEntryId { get; set; } = string.Empty;
    public int TicketMinutesBooked { get; set; }
    public long TicketSecondsBooked { get; set; }
    public bool IsPinned { get; set; }
    public bool IsZnunyAssigned { get; set; } = true;
    public bool IsZnunyTask => (Tags ?? string.Empty).Contains("ZnunyTicketID:", StringComparison.OrdinalIgnoreCase);
    public bool IsOperationallyVisible => !IsZnunyTask || IsZnunyAssigned;
    // Derived presentation state; never persisted.
    public string CurrentListBadgeText { get; set; } = string.Empty;
    public bool ShowCurrentListBadge => !string.IsNullOrWhiteSpace(CurrentListBadgeText);
    public bool IsCurrentListBadgePlanned => string.Equals(CurrentListBadgeText, "Geplant", StringComparison.Ordinal);
    public string CurrentListTicketNumber { get; set; } = string.Empty;
    public string CurrentListDisplayTitle { get; set; } = string.Empty;
    public string CurrentListPreviewText { get; set; } = string.Empty;
    public bool ShowCurrentListTicketNumber => !string.IsNullOrWhiteSpace(CurrentListTicketNumber);
    public bool ShowCurrentListPreview => !string.IsNullOrWhiteSpace(CurrentListPreviewText);
    public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
    public DateTime? TicketCreatedUtc { get; set; }
    public DateTime? TicketChangedUtc { get; set; }
    public DateTime LocalActivityUtc { get; set; } = DateTime.UtcNow;
}
