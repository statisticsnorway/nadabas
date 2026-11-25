using Nadabas.PackageUpdate;

namespace Nadabas.Setup.Tests;

public sealed class RestartAcknowledgementTests
{
    [Fact]
    public async Task WaitsUntilTheHelperConfirmsReadiness()
    {
        var completion = new TaskCompletionSource<string?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var waiting = RestartAcknowledgement.RequireReady(completion.Task, TimeSpan.FromSeconds(5));
        Assert.False(waiting.IsCompleted);
        completion.SetResult("restart-helper-ready");
        await waiting;
    }

    [Theory]
    [InlineData(null)]
    [InlineData("failed")]
    public async Task MissingOrWrongAcknowledgementIsRejected(string? acknowledgement)
    {
        await Assert.ThrowsAsync<IOException>(() => RestartAcknowledgement.RequireReady(
            Task.FromResult(acknowledgement), TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public async Task UnresponsiveHelperDoesNotBecomeReady()
    {
        var completion = new TaskCompletionSource<string?>();
        await Assert.ThrowsAsync<TimeoutException>(() => RestartAcknowledgement.RequireReady(completion.Task, TimeSpan.Zero));
    }
}
