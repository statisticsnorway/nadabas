using System.Diagnostics;

namespace Nadabas.Versioning;

// Separate process: never terminate Excel, suppress save prompts, or reopen
// workbooks. A bounded wait also limits a cancelled Quit request's lifetime.
public static class ExcelRestart
{
    public static async Task Run(string root, int processId, long startTicks)
    {
        if (!OperatingSystem.IsWindows()) throw new IOException("Excel restart requires Windows.");
        var store = new VersionStore(root);
        var ready = store.GetReadyUpdate() ?? throw new IOException("No verified update is ready.");
        using var parent = Process.GetProcessById(processId);
        var executable = parent.MainModule?.FileName;
        if (parent.StartTime.ToUniversalTime().Ticks != startTicks ||
            !string.Equals(Path.GetFileName(executable), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
            throw new IOException("The original Excel process could not be identified.");
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
        _ = parent.SafeHandle; // Acquire the original process handle before acknowledging.
        Console.WriteLine("restart-helper-ready");
        await AfterExit(parent.WaitForExitAsync, () => store.GetReadyUpdate() == ready,
            AnyExcelRunning, () => Process.Start(new ProcessStartInfo(executable!) { UseShellExecute = true, Arguments = "/x" }), timeout.Token);
    }

    private static bool AnyExcelRunning()
    {
        var others = Process.GetProcessesByName("EXCEL");
        try { return others.Length != 0; }
        finally { foreach (var other in others) other.Dispose(); }
    }

    internal static async Task AfterExit(Func<CancellationToken, Task> wait, Func<bool> sameUpdate,
        Func<bool> excelRunning, Action launch, CancellationToken cancellation)
    {
        await wait(cancellation);
        cancellation.ThrowIfCancellationRequested();
        if (excelRunning()) throw new IOException("Another Excel process is still running.");
        if (!sameUpdate()) throw new IOException("The pending update changed; automatic restart cancelled.");
        launch();
    }
}
