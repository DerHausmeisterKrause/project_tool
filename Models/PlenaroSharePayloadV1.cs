namespace TaskTool.Models;

public sealed class PlenaroSharePayloadV1
{
    public int Version { get; set; } = 1;
    public string TaskShareId { get; set; } = string.Empty;
    public string SegmentShareId { get; set; } = string.Empty;
    public string OriginClientInstanceId { get; set; } = string.Empty;
    public string OriginTaskId { get; set; } = string.Empty;
    public long OriginSegmentId { get; set; }
    public string TaskTitle { get; set; } = string.Empty;
    public string TaskDescription { get; set; } = string.Empty;
    public string TicketId { get; set; } = string.Empty;
    public string TicketNumber { get; set; } = string.Empty;
    public string TicketUrl { get; set; } = string.Empty;
    public string TicketState { get; set; } = string.Empty;
    public DateTime SegmentStartLocal { get; set; }
    public DateTime SegmentEndLocal { get; set; }
    public string SegmentNote { get; set; } = string.Empty;
    public DateTime GeneratedUtc { get; set; }
}
