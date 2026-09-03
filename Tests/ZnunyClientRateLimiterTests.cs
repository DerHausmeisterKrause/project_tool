using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunyClientRateLimiterTests
{
    [Fact]
    public async Task RequestsObserveMinimumSpacing()
    {
        var time = new AutoAdvancingTimeProvider();
        using var limiter = new ZnunyClientRateLimiter(time, minimumSpacing: TimeSpan.FromMilliseconds(40), windowLimit: 100);
        var times = new List<DateTimeOffset>();
        for (var i = 0; i < 3; i++) { await limiter.WaitAsync(); times.Add(time.GetUtcNow()); }
        Assert.All(times.Zip(times.Skip(1)), pair => Assert.Equal(TimeSpan.FromMilliseconds(40), pair.Second - pair.First));
    }

    [Fact]
    public async Task FullWindowWaitsInsteadOfRejecting()
    {
        using var limiter = new ZnunyClientRateLimiter(minimumSpacing: TimeSpan.Zero,
            windowLimit: 3, windowLength: TimeSpan.FromMilliseconds(80));
        for (var i = 0; i < 4; i++) await limiter.WaitAsync();
        Assert.True(limiter.WaitCount >= 1);
        Assert.True(limiter.TotalDelayMilliseconds >= 70);
    }

    private sealed class AutoAdvancingTimeProvider : TimeProvider
    {
        private DateTimeOffset _utcNow = new(2026, 1, 1, 0, 0, 0, TimeSpan.Zero);
        public override DateTimeOffset GetUtcNow() { lock (this) return _utcNow; }
        public override ITimer CreateTimer(TimerCallback callback, object? state, TimeSpan dueTime, TimeSpan period)
        {
            var timer = new ImmediateTimer();
            ThreadPool.QueueUserWorkItem(_ =>
            {
                lock (this) _utcNow += dueTime;
                if (!timer.IsDisposed) callback(state);
            });
            return timer;
        }

        private sealed class ImmediateTimer : ITimer
        {
            public bool IsDisposed { get; private set; }
            public bool Change(TimeSpan dueTime, TimeSpan period) => !IsDisposed;
            public void Dispose() => IsDisposed = true;
            public ValueTask DisposeAsync() { Dispose(); return ValueTask.CompletedTask; }
        }
    }
}
