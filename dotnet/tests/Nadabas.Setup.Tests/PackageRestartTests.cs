using Nadabas.PackageUpdate;

namespace Nadabas.Setup.Tests;

public sealed class PackageRestartTests
{
    [Fact]
    public async Task CancelledExcelShutdownNeverInstallsOrReopens()
    {
        var applied = false;
        var reopened = false;
        using var timeout = new CancellationTokenSource();
        timeout.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => PackageRestart.AfterExit(
            token => Task.FromCanceled(token), () => false,
            () => { applied = true; return Task.FromResult(true); }, () => reopened = true, timeout.Token));
        Assert.False(applied);
        Assert.False(reopened);
    }

    [Fact]
    public async Task AnotherExcelProcessDefersInstallation()
    {
        var applied = false;
        await PackageRestart.AfterExit(_ => Task.CompletedTask, () => true,
            () => { applied = true; return Task.FromResult(true); }, () => throw new InvalidOperationException(), default);
        Assert.False(applied);
    }

    [Fact]
    public async Task ExcelReopensOnlyAfterSuccessfulInstallation()
    {
        var events = new List<string>();
        await PackageRestart.AfterExit(_ => { events.Add("exit"); return Task.CompletedTask; }, () => false,
            () => { events.Add("install"); return Task.FromResult(true); }, () => events.Add("reopen"), default);
        Assert.Equal(new[] { "exit", "install", "reopen" }, events);
    }

    [Fact]
    public async Task RejectedOrFailedInstallationDoesNotReopenExcel()
    {
        var reopened = false;
        await PackageRestart.AfterExit(_ => Task.CompletedTask, () => false, () => Task.FromResult(false), () => reopened = true, default);
        await Assert.ThrowsAsync<IOException>(() => PackageRestart.AfterExit(_ => Task.CompletedTask, () => false,
            () => throw new IOException("Policy changed"), () => reopened = true, default));
        Assert.False(reopened);
    }

    [Fact]
    public async Task NewExcelProcessDuringInstallPreventsDuplicateLaunch()
    {
        var running = false;
        await PackageRestart.AfterExit(_ => Task.CompletedTask, () => running,
            () => { running = true; return Task.FromResult(true); }, () => throw new InvalidOperationException(), default);
    }
}
