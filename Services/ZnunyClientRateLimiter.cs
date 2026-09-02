using System.Diagnostics;

namespace TaskTool.Services;

/// <summary>Per-process limiter for the single central Znuny HTTP path.</summary>
public sealed class ZnunyClientRateLimiter : IDisposable
{
    public static readonly TimeSpan DefaultMinimumSpacing = TimeSpan.FromMilliseconds(150);
    public static readonly TimeSpan DefaultWindow = TimeSpan.FromSeconds(60);
    public const int DefaultWindowLimit = 300;

    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly Queue<DateTimeOffset> _window = new();
    private readonly TimeProvider _time;
    private readonly TimeSpan _minimumSpacing;
    private readonly TimeSpan _windowLength;
    private readonly int _windowLimit;
    private DateTimeOffset? _lastRequest;
    private long _delayMilliseconds;
    private long _waitCount;

    public ZnunyClientRateLimiter(TimeProvider? timeProvider = null, TimeSpan? minimumSpacing = null,
        int windowLimit = DefaultWindowLimit, TimeSpan? windowLength = null)
    {
        _time = timeProvider ?? TimeProvider.System;
        _minimumSpacing = minimumSpacing ?? DefaultMinimumSpacing;
        _windowLimit = windowLimit;
        _windowLength = windowLength ?? DefaultWindow;
    }

    public long TotalDelayMilliseconds => Interlocked.Read(ref _delayMilliseconds);
    public long WaitCount => Interlocked.Read(ref _waitCount);

    public async ValueTask WaitAsync(CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            while (true)
            {
                var now = _time.GetUtcNow();
                while (_window.TryPeek(out var oldest) && now - oldest >= _windowLength) _window.Dequeue();
                var spacingWait = _lastRequest is { } last ? _minimumSpacing - (now - last) : TimeSpan.Zero;
                var windowWait = _window.Count >= _windowLimit ? _window.Peek() + _windowLength - now : TimeSpan.Zero;
                var delay = spacingWait > windowWait ? spacingWait : windowWait;
                if (delay <= TimeSpan.Zero)
                {
                    _lastRequest = now;
                    _window.Enqueue(now);
                    return;
                }
                Interlocked.Increment(ref _waitCount);
                Interlocked.Add(ref _delayMilliseconds, (long)Math.Ceiling(delay.TotalMilliseconds));
                await Task.Delay(delay, _time, cancellationToken).ConfigureAwait(false);
            }
        }
        finally { _gate.Release(); }
    }

    public void Dispose() => _gate.Dispose();
}
