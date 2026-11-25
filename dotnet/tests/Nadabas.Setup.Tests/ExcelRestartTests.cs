using Nadabas.Versioning;

namespace Nadabas.Setup.Tests;

public sealed class ExcelRestartTests
{
    [Fact]
    public async Task RelaunchWaitsForOriginalProcessExit()
    {
        var exit = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var launched = false;
        var pending = ExcelRestart.AfterExit(_ => exit.Task, () => true, () => false, () => launched = true, CancellationToken.None);
        Assert.False(launched);
        exit.SetResult();
        await pending;
        Assert.True(launched);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public async Task ChangedUpdateOrOtherExcelPreventsRelaunch(bool sameUpdate, bool otherExcel)
    {
        var launched = false;
        await Assert.ThrowsAsync<IOException>(() => ExcelRestart.AfterExit(_ => Task.CompletedTask,
            () => sameUpdate, () => otherExcel, () => launched = true, CancellationToken.None));
        Assert.False(launched);
    }

    [Fact]
    public async Task CancelledShutdownNeverLaunchesAnotherExcel()
    {
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        var launched = false;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelRestart.AfterExit(
            token => Task.Delay(Timeout.Infinite, token), () => true, () => false, () => launched = true, cancelled.Token));
        Assert.False(launched);
    }
}
