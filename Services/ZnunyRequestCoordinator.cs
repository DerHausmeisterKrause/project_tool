namespace TaskTool.Services;

/// <summary>Serializes all remote Znuny pipelines without making nested core calls reacquire the gate.</summary>
public sealed class ZnunyRequestCoordinator : IDisposable
{
    private readonly SemaphoreSlim _gate = new(1, 1);
    public Task<bool> WaitAsync(int millisecondsTimeout) => _gate.WaitAsync(millisecondsTimeout);
    public Task<bool> WaitAsync(TimeSpan timeout) => _gate.WaitAsync(timeout);
    public void Release() => _gate.Release();
    public void Dispose() => _gate.Dispose();
}
