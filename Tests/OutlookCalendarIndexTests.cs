using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class OutlookCalendarIndexTests
{
    [Fact]
    public void DayIndexPreservesAllOverlapCasesAndOrdering()
    {
        var day = new DateTime(2026, 10, 12);
        var beforeIntoRange = Event("before", day.AddHours(8), day.AddHours(9.5));
        var inside = Event("inside", day.AddHours(10), day.AddHours(11));
        var insidePastRange = Event("past", day.AddHours(11.5), day.AddHours(14));
        var allDay = Event("all-day", day, day.AddDays(1), allDay: true);
        var recurringInstance = Event("recurring", day.AddHours(10.5), day.AddHours(10.75));
        recurringInstance.IsRecurring = true; recurringInstance.IsInstance = true;
        var outside = Event("outside", day.AddDays(-2), day.AddDays(-2).AddHours(1));
        var duplicateIdentity = Event("inside", inside.StartLocal, inside.EndLocal);

        var index = OutlookCalendarService.BuildDayIndex([outside, insidePastRange, inside, beforeIntoRange, allDay, recurringInstance, duplicateIdentity]);
        var result = OutlookCalendarService.FilterIndexedEvents(index, day.AddHours(9), day.AddHours(12));

        Assert.Equal(["all-day", "before", "inside", "inside", "recurring", "past"], result.Select(item => item.Id));
        Assert.DoesNotContain(result, item => item.Id == "outside");
    }

    private static OutlookCalendarEvent Event(string id, DateTime start, DateTime end, bool allDay = false)
        => new() { Id = id, EntryId = id, StartLocal = start, EndLocal = end, IsAllDay = allDay };
}
