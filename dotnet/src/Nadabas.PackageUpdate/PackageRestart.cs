namespace Nadabas.PackageUpdate;

internal static class PackageRestart
{
    internal static async Task AfterExit(Func<CancellationToken, Task> wait,
        Func<bool> excelRunning, Func<Task<bool>> apply, Action reopen, CancellationToken cancellation)
    {
        await wait(cancellation).ConfigureAwait(false);
        cancellation.ThrowIfCancellationRequested();
        if (excelRunning()) return;
        if (!await apply().ConfigureAwait(false)) return;
        if (!excelRunning()) reopen();
    }
}
