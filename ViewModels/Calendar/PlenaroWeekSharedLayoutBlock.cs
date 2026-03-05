namespace TaskTool.ViewModels;

internal sealed class PlenaroWeekSharedLayoutBlock
{
    public DateTime Start { get; }
    public DateTime End { get; }
    public Action<int, int> Assign { get; }
    public int Column { get; set; }

    public PlenaroWeekSharedLayoutBlock(DateTime start, DateTime end, Action<int, int> assign)
    {
        Start = start;
        End = end;
        Assign = assign;
    }
}
