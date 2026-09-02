using System.Security.Cryptography;
using System.Text;

namespace TaskTool.Models;

public static class ZnunySyncSchedulerPolicy
{
    public static TimeSpan StartupCandidateDelay(string clientInstanceId) => StableOffset(clientInstanceId, "startup-candidate", 2, 12);
    public static TimeSpan StartupFullDelay(string clientInstanceId) => StableOffset(clientInstanceId, "startup-full", 5, 30);
    public static TimeSpan BusyRetryDelay(string clientInstanceId, string pipeline) => StableOffset(clientInstanceId, "busy-" + pipeline, 10, 30);

    public static DateTimeOffset NextSlot(DateTimeOffset anchor, DateTimeOffset now, TimeSpan interval)
    {
        if (now < anchor) return anchor;
        var slots = (long)Math.Floor((now - anchor).Ticks / (double)interval.Ticks) + 1;
        return anchor + TimeSpan.FromTicks(checked(slots * interval.Ticks));
    }

    private static TimeSpan StableOffset(string id, string salt, int minimumSeconds, int maximumSeconds)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(id + ":" + salt));
        var range = maximumSeconds - minimumSeconds + 1;
        return TimeSpan.FromSeconds(minimumSeconds + BitConverter.ToUInt32(hash, 0) % range);
    }
}
