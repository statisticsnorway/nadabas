using System.Diagnostics;
using System.Windows.Forms;
using ExcelDna.Integration;
using Nadabas.PackageUpdate;

namespace Nadabas.ExcelStartup;

public sealed class Startup : IExcelAddIn
{
    private StartupPayload? payload;

    public void AutoOpen() => ExcelAsyncUtil.QueueAsMacro(Start);

    private void Start()
    {
        try
        {
            if (payload is not null) return;
            // Do not change runtime selection for another already-running modern add-in.
            using (var process = Process.GetCurrentProcess())
                foreach (ProcessModule module in process.Modules)
                    if (string.Equals(module.ModuleName, "coreclr.dll", StringComparison.OrdinalIgnoreCase))
                        throw new IOException("A modern .NET runtime is already loaded. This engineering prototype cannot share it yet.");
            var root = Path.GetDirectoryName(ExcelDnaUtil.XllPath)!;
            using var resource = typeof(Startup).Assembly.GetManifestResourceStream("Nadabas.StartupManifest")
                ?? throw new IOException("Missing embedded startup manifest.");
            using var reader = new StreamReader(resource);
            payload = StartupPayload.Verify(root, StartupPayload.Architecture(IntPtr.Size),
                reader.ReadToEnd().Split('\n').Select(line => line.TrimEnd('\r')), PublisherTrust.Verify);
            dynamic excel = ExcelDnaUtil.Application;
            if (!payload.Register(path => (bool)excel.RegisterXLL(path)))
                throw new IOException("Excel did not activate the verified modern loader.");
        }
        catch (Exception ex)
        {
            payload?.Dispose();
            payload = null;
            MessageBox.Show("NADABAS startup prototype could not start.\n\n" + ex.Message,
                "NADABAS startup prototype", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    // Normal shutdown releases process handles; an explicit removal also releases leases.
    public void AutoClose() { payload?.Dispose(); payload = null; }
}
