using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class SegmentAvailabilitySlotTests
{
    private static readonly DateTime Day = new(2026, 8, 26);

    [Fact]
    public void AppointmentFromNineToNineThirtyOccupiesExactlyTwoSlots()
    {
        var slots = Enumerable.Range(0, 48).Select(index => new SegmentAvailabilitySlot
        {
            Start = Day.AddHours(6).AddMinutes(index * 15),
            End = Day.AddHours(6).AddMinutes((index + 1) * 15)
        });

        Assert.Equal(2, slots.Count(slot => slot.Overlaps(Day.AddHours(9), Day.AddHours(9.5))));
    }

    [Fact]
    public void BoundaryTouchDoesNotOverlap()
    {
        var slot = new SegmentAvailabilitySlot { Start = Day.AddHours(9), End = Day.AddHours(9.25) };

        Assert.False(slot.Overlaps(Day.AddHours(8), Day.AddHours(9)));
        Assert.False(slot.Overlaps(Day.AddHours(9.25), Day.AddHours(10)));
    }

    [Fact]
    public void UnknownSlotIsNotReportedAsAvailable()
    {
        var slot = new SegmentAvailabilitySlot { Start = Day, End = Day.AddMinutes(15), IsUnknown = true };
        Assert.False(slot.IsAvailable);
    }
}
