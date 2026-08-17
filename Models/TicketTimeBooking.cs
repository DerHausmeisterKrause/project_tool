namespace TaskTool.Models;

public class TicketTimeBooking
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public Guid TaskId { get; set; }
    public string TicketId { get; set; } = string.Empty;
    public string TicketNumber { get; set; } = string.Empty;
    public string BookingId { get; set; } = Guid.NewGuid().ToString("D");
    public string ArticleId { get; set; } = string.Empty;
    public decimal Minutes { get; set; }
    public decimal BookedMinutes { get; set; }
    public long SourceSeconds { get; set; }
    public DateTime BookedAtUtc { get; set; } = DateTime.UtcNow;
    public string ShortDescription { get; set; } = string.Empty;
    public string CostCenter { get; set; } = string.Empty;
    public string Order { get; set; } = string.Empty;
    public string Status { get; set; } = "Pending";

    public string DisplayTime => BookedAtUtc.ToLocalTime().ToString("dd.MM.yyyy HH:mm");
    public string DisplayMinutes => $"{Minutes:0.##} Min. erfasst → {BookedMinutes:0.##} Min. gebucht";
    public string StatusText => Status switch
    {
        "Succeeded" => "✓ Gebucht",
        "Pending" => "⏳ Unklar / Pending",
        _ => "✕ Fehlgeschlagen"
    };
    public bool CanCheckStatus => Status == "Pending";
    public bool CanRetry => Status == "Failed";
}
