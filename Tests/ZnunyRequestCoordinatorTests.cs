using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunyRequestCoordinatorTests
{
    [Fact]
    public async Task BackgroundAndManualPipelinesCannotRunConcurrently()
    {
        using var coordinator = new ZnunyRequestCoordinator();
        Assert.True(await coordinator.WaitAsync(0));
        Assert.False(await coordinator.WaitAsync(TimeSpan.FromMilliseconds(20)));
        coordinator.Release();
        Assert.True(await coordinator.WaitAsync(0));
        coordinator.Release();
    }
}
