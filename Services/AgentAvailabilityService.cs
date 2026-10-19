using System.Collections.Concurrent;
using TaskTool.Models;

namespace TaskTool.Services;

public interface IAgentAvailabilityService
{
    Task<AgentAvailabilitySnapshot> GetAvailabilityAsync(ZnunyAgent agent, CancellationToken token = default);
}

public sealed class AgentAvailabilityService : IAgentAvailabilityService
{
    private static readonly TimeSpan CacheDuration = TimeSpan.FromMinutes(10);
    private readonly Func<ZnunyAgent, Task<OutlookFreeBusySnapshot?>> _loadFreeBusy;
    private readonly Func<int, Task<int?>> _loadOpenTickets;
    private readonly Func<DateTime> _utcNow;
    private readonly Func<DateTime> _localNow;
    private readonly ConcurrentDictionary<int, CacheEntry> _cache = new();
    private readonly ConcurrentDictionary<int, Task<AgentAvailabilitySnapshot>> _inflight = new();

    public AgentAvailabilityService(OutlookInteropService outlook, TicketSystemService tickets)
        : this(
            agent => Task.Run(() =>
            {
                var start = DateTime.Today;
                return outlook.TryGetRecipientFreeBusy(agent, start, start.AddDays(14));
            }),
            userId => tickets.TryGetAgentOpenTicketCountAsync(userId),
            () => DateTime.UtcNow,
            () => DateTime.Now)
    {
    }

    internal AgentAvailabilityService(
        Func<ZnunyAgent, Task<OutlookFreeBusySnapshot?>> loadFreeBusy,
        Func<int, Task<int?>> loadOpenTickets,
        Func<DateTime>? utcNow = null,
        Func<DateTime>? localNow = null)
    {
        _loadFreeBusy = loadFreeBusy;
        _loadOpenTickets = loadOpenTickets;
        _utcNow = utcNow ?? (() => DateTime.UtcNow);
        _localNow = localNow ?? (() => DateTime.Now);
    }

    public async Task<AgentAvailabilitySnapshot> GetAvailabilityAsync(ZnunyAgent agent, CancellationToken token = default)
    {
        var now = _utcNow();
        if (_cache.TryGetValue(agent.UserId, out var cached) && cached.ExpiresUtc > now)
            return cached.Snapshot;

        var request = _inflight.GetOrAdd(agent.UserId, _ => CalculateAsync(agent));
        try { return await request.WaitAsync(token).ConfigureAwait(false); }
        finally
        {
            if (request.IsCompleted && _inflight.TryGetValue(agent.UserId, out var current) && ReferenceEquals(current, request))
                _inflight.TryRemove(agent.UserId, out _);
        }
    }

    private async Task<AgentAvailabilitySnapshot> CalculateAsync(ZnunyAgent agent)
    {
        var freeBusy = await _loadFreeBusy(agent).ConfigureAwait(false);
        var calculated = _utcNow();
        if (freeBusy == null || !TryCalculateCalendarLoad(freeBusy, _localNow(), out var calendarLoad))
            return Cache(agent.UserId, new(AgentAvailabilityLevel.Unknown, calculated));

        int? ticketCount = null;
        try { ticketCount = await _loadOpenTickets(agent.UserId).ConfigureAwait(false); }
        catch { /* Optional source: Outlook remains authoritative. */ }

        return Cache(agent.UserId, new(Classify(calendarLoad, ticketCount), calculated));
    }

    private AgentAvailabilitySnapshot Cache(int userId, AgentAvailabilitySnapshot snapshot)
    {
        _cache[userId] = new(snapshot, _utcNow() + CacheDuration);
        return snapshot;
    }

    internal static AgentAvailabilityLevel Classify(double? calendarLoad, int? openTicketCount)
    {
        if (calendarLoad == null) return AgentAvailabilityLevel.Unknown;
        var score = openTicketCount.HasValue
            ? calendarLoad.Value * .8 + Math.Clamp(openTicketCount.Value / 20d, 0, 1) * .2
            : calendarLoad.Value;
        return score < .55 ? AgentAvailabilityLevel.Green : score < .8 ? AgentAvailabilityLevel.Orange : AgentAvailabilityLevel.Red;
    }

    internal static bool TryCalculateCalendarLoad(OutlookFreeBusySnapshot snapshot, DateTime nowLocal, out double load)
    {
        load = 0;
        if (snapshot.IntervalMinutes <= 0 || string.IsNullOrEmpty(snapshot.Slots)) return false;
        var end = nowLocal.AddDays(14);
        double occupied = 0;
        var considered = 0;
        for (var i = 0; i < snapshot.Slots.Length; i++)
        {
            var start = snapshot.StartLocal.AddMinutes((long)i * snapshot.IntervalMinutes);
            var slotEnd = start.AddMinutes(snapshot.IntervalMinutes);
            if (start >= end) break;
            if (slotEnd <= nowLocal || start.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) continue;
            if (start.TimeOfDay < TimeSpan.FromHours(8) || slotEnd.TimeOfDay > TimeSpan.FromHours(17)) continue;
            var weight = snapshot.Slots[i] switch
            {
                '0' or '4' => 0d,
                '1' => .5d,
                '2' or '3' => 1d,
                _ => double.NaN
            };
            if (double.IsNaN(weight)) return false;
            occupied += weight;
            considered++;
        }
        if (considered == 0) return false;
        load = occupied / considered;
        return true;
    }

    private sealed record CacheEntry(AgentAvailabilitySnapshot Snapshot, DateTime ExpiresUtc);
}
