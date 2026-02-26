namespace TaskTool.Models;

public class TaskSegment
{
    private DateTime _startLocal;
    private DateTime _endLocal;

    public long Id { get; set; }
    public Guid TaskId { get; set; }

    public DateTime StartLocal
    {
        get => _startLocal;
        set => _startLocal = value;
    }

    public DateTime EndLocal
    {
        get => _endLocal;
        set => _endLocal = value;
    }

    public DateTime SegmentDate
    {
        get => StartLocal.Date;
        set
        {
            StartLocal = value.Date + StartLocal.TimeOfDay;
            EndLocal = value.Date + EndLocal.TimeOfDay;
        }
    }

    public string StartTimeText
    {
        get => StartLocal.ToString("HH:mm");
        set
        {
            if (TimeSpan.TryParse(value, out var time))
                StartLocal = StartLocal.Date + time;
        }
    }

    public string EndTimeText
    {
        get => EndLocal.ToString("HH:mm");
        set
        {
            if (TimeSpan.TryParse(value, out var time))
                EndLocal = EndLocal.Date + time;
        }
    }


    public string ValidationHint
    {
        get
        {
            if (StartLocal == default) return "Datum muss gesetzt sein.";
            if (EndLocal == default) return "Endzeit darf nicht leer sein.";
            if (StartLocal >= EndLocal) return "Startzeit muss vor Endzeit liegen.";
            return string.Empty;
        }
    }

    public bool IsValid => string.IsNullOrWhiteSpace(ValidationHint);

    public int PlannedMinutes { get; set; }
    public string Note { get; set; } = string.Empty;
    public string OutlookEntryId { get; set; } = string.Empty;
    public string OutlookStatus { get; set; } = string.Empty;
}
