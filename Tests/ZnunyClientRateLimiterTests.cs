using System.Diagnostics;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunyClientRateLimiterTests
{
    [Fact]
    public async Task RequestsObserveMinimumSpacing()
    {
        using var limiter = new ZnunyClientRateLimiter(minimumSpacing: TimeSpan.FromMilliseconds(40), windowLimit: 100);
        var times = new List<long>();
        var stopwatch = Stopwatch.StartNew();
        for (var i = 0; i < 3; i++) { await limiter.WaitAsync(); times.Add(stopwatch.ElapsedMilliseconds); }
        Assert.All(times.Zip(times.Skip(1)), pair => Assert.True(pair.Second - pair.First >= 35));
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
}
