using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Nadabas.Versioning;
using Nadabas.PackageUpdate;

namespace Nadabas.ExcelBootstrap;

internal sealed class RestartBlockedException(string message) : IOException(message);

internal static class RestartUpdate
{
    internal static void Start(string root, string language, string? platformId)
    {
        if (platformId is null) _ = new VersionStore(root).GetReadyUpdate() ?? throw new IOException("No verified update is ready.");
        var processes = Process.GetProcessesByName("EXCEL");
        try
        {
            if (processes.Length != 1 || processes[0].Id != Environment.ProcessId)
                throw new RestartBlockedException(UpdateLanguage.Text(language, "otherExcel"));
        }
        finally { foreach (var process in processes) process.Dispose(); }
        dynamic excel = ExcelDnaUtil.Application;
        dynamic books = excel.Workbooks;
        try
        {
            for (var i = 1; i <= (int)books.Count; i++)
            {
                dynamic book = books.Item(i);
                try
                {
                    if (!(bool)book.IsAddin && !(bool)book.Saved)
                        throw new RestartBlockedException(UpdateLanguage.Text(language, "unsaved"));
                }
                finally { Marshal.ReleaseComObject(book); }
            }
        }
        finally { Marshal.ReleaseComObject(books); }
        using var current = Process.GetCurrentProcess();
        using var helper = platformId is null ? StartSlotHelper(root, current) : PlatformHandoff.LaunchWaiter(root, platformId, current);
        try
        {
            var acknowledgement = helper.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult();
            if (acknowledgement != "restart-helper-ready")
            {
                var reason = helper.StartInfo.RedirectStandardError
                    ? helper.StandardError.ReadToEndAsync().WaitAsync(TimeSpan.FromSeconds(2)).GetAwaiter().GetResult().Trim()
                    : string.Empty;
                throw new IOException("Restart helper did not become ready. " + reason);
            }
        }
        catch
        {
            if (!helper.HasExited) helper.Kill(); // Only our failed helper, never Excel.
            throw;
        }
        // Preserve ordinary save/cancellation behavior even if another add-in
        // modifies a workbook during shutdown. The helper only waits for exit.
        excel.DisplayAlerts = true;
        excel.Quit();
    }

    private static Process StartSlotHelper(string root, Process current)
    {
        var command = new ProcessStartInfo(Path.Combine(root, "tools", "Nadabas.VersionStore.exe"))
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true
        };
        command.ArgumentList.Add("restart-excel");
        command.ArgumentList.Add(root);
        command.ArgumentList.Add(current.Id.ToString(System.Globalization.CultureInfo.InvariantCulture));
        command.ArgumentList.Add(current.StartTime.ToUniversalTime().Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture));
        return Process.Start(command) ?? throw new IOException("Could not start the restart helper.");
    }
}
