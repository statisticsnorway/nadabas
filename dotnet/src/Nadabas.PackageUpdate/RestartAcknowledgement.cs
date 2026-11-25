using System.Diagnostics;

namespace Nadabas.PackageUpdate;

internal static class RestartAcknowledgement
{
    internal static void Wait(Process process) =>
        RequireReady(process.StandardOutput.ReadLineAsync(), TimeSpan.FromSeconds(30)).GetAwaiter().GetResult();

    internal static async Task RequireReady(Task<string?> acknowledgement, TimeSpan timeout)
    {
        if (await acknowledgement.WaitAsync(timeout) != "restart-helper-ready")
            throw new IOException("Restart helper did not acquire its runtime and become ready. Check the update log.");
    }
}
