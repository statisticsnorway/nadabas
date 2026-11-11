using System.Text.Json;

namespace Nadabas.Setup;

public static class Command
{
    public static int Run(string[] args, TextWriter output, TextWriter error)
    {
        if (args.Length == 1 && args[0] is "--help" or "-h")
        {
            output.WriteLine("NADABAS Setup & Diagnostics (diagnostics, planning and isolated file rehearsal)");
            output.WriteLine("Usage: Nadabas.Setup diagnose [--addin-directory <absolute-path>]");
            output.WriteLine("       Nadabas.Setup plan-legacy-migration [--addin-directory <absolute-path>] (read-only)");
            output.WriteLine("       Nadabas.Setup plan-install --source <absolute-xlam-path> --addin-directory <absolute-path>");
            output.WriteLine("       Nadabas.Setup rehearse-install --source <absolute-xlam-path> --output <absolute-test-parent>");
            output.WriteLine("       Nadabas.Setup rehearse-update --source <absolute-xlam-path> --output <absolute-test-parent>");
            output.WriteLine("       Nadabas.Setup rehearse-update-live --source <absolute-xlam-path> --output <absolute-test-parent>");
            output.WriteLine("       Nadabas.Setup recover-test --workspace <absolute-generated-test-directory>");
            output.WriteLine("       Nadabas.Setup excel-startup-test --source <absolute-xlam-path> --output <absolute-test-parent> [--interactive]");
            output.WriteLine("Output: JSON schemaVersion 1. Exit 0: report complete; 1: incomplete; 2: invalid arguments.");
            return 0;
        }

        if (args.Length > 0 && args[0] == "plan-legacy-migration")
        {
            if (args.Length != 1 && !(args.Length == 3 && args[1] == "--addin-directory" && Path.IsPathFullyQualified(args[2])))
            { error.WriteLine("Use plan-legacy-migration [--addin-directory <absolute-path>]."); return 2; }
            var legacyDirectory = args.Length == 3 ? args[2] : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "AddIns");
            WriteJson(LegacyMigration.Collect(legacyDirectory), output);
            return 0;
        }
        if (args.Length > 0 && args[0] == "plan-install")
            return PlanInstall(args, output, error);
        if (args.Length > 0 && args[0] == "excel-startup-test")
        {
            if (args.Length is not (5 or 6) || args[1] != "--source" || args[3] != "--output" ||
                !Path.IsPathFullyQualified(args[2]) || !Path.IsPathFullyQualified(args[4]) || (args.Length == 6 && args[5] != "--interactive"))
            {
                error.WriteLine("Use --help for Excel startup test syntax; paths must be absolute.");
                return 2;
            }
            if (!OperatingSystem.IsWindows()) { error.WriteLine("Desktop Excel startup testing requires Windows."); return 1; }
            try
            {
                var startupReport = ExcelStartupTest.Run(Path.GetFullPath(args[2]), Path.GetFullPath(args[4]), args.Length == 6, error);
                WriteJson(startupReport, output);
                return startupReport.Status == "activation-verified-ribbon-pending" ? 0 : 1;
            }
            catch (Exception ex) { error.WriteLine("Excel startup test stopped: " + ex.Message); return 1; }
        }
        if (args.Length > 0 && args[0] is "rehearse-install" or "rehearse-update" or "rehearse-update-live" or "recover-test")
            return Rehearse(args, output, error);

        if (!(args.Length == 1 && args[0] == "diagnose") &&
            !(args.Length == 3 && args[0] == "diagnose" && args[1] == "--addin-directory"))
        {
            error.WriteLine("Use: Nadabas.Setup diagnose [--addin-directory <absolute-path>]");
            return 2;
        }

        string directory;
        try
        {
            directory = args.Length == 3
                ? args[2]
                : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "AddIns");
            if (!Path.IsPathFullyQualified(directory))
                throw new ArgumentException();
            directory = Path.GetFullPath(directory);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            error.WriteLine("The add-in directory must be a valid absolute path.");
            return 2;
        }

        var report = Diagnostics.Collect(directory);
        output.WriteLine(JsonSerializer.Serialize(report, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        }));
        return report.Complete ? 0 : 1;
    }

    private static int Rehearse(string[] args, TextWriter output, TextWriter error)
    {
        try
        {
            if (args[0] == "recover-test" && args.Length == 3 && args[1] == "--workspace" && Path.IsPathFullyQualified(args[2]))
            {
                var result = new TestFileTransaction(args[2]).Recover();
                WriteJson(result, output);
                return 0;
            }
            if (args[0] is "rehearse-install" or "rehearse-update" or "rehearse-update-live" && args.Length == 5 && args[1] == "--source" && args[3] == "--output" &&
                Path.IsPathFullyQualified(args[2]) && Path.IsPathFullyQualified(args[4]))
            {
                if (args[0] == "rehearse-update-live")
                {
                    if (!OperatingSystem.IsWindows()) { error.WriteLine("Live Excel testing requires Windows."); return 1; }
                    var liveReport = LiveUpdateRehearsal.Run(Path.GetFullPath(args[2]), Path.GetFullPath(args[4]));
                    WriteJson(liveReport, output);
                    return liveReport.Passed ? 0 : 1;
                }
                if (args[0] == "rehearse-update")
                {
                    var updateReport = UpdateRehearsal.Run(Path.GetFullPath(args[2]), Path.GetFullPath(args[4]));
                    WriteJson(updateReport, output);
                    return updateReport.Passed ? 0 : 1;
                }
                var report = InstallationRehearsal.Run(Path.GetFullPath(args[2]), Path.GetFullPath(args[4]));
                WriteJson(report, output);
                return report.Passed ? 0 : 1;
            }
            error.WriteLine("Use --help for rehearsal/recovery syntax. Paths must be absolute.");
            return 2;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException or JsonException or ArgumentException or NotSupportedException)
        {
            error.WriteLine("Test operation stopped: " + ex.Message);
            return 1;
        }
    }

    private static void WriteJson<T>(T value, TextWriter output) => output.WriteLine(JsonSerializer.Serialize(value,
        new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true }));

    private static int PlanInstall(string[] args, TextWriter output, TextWriter error)
    {
        if (args.Length != 5 || args[1] != "--source" || args[3] != "--addin-directory")
        {
            error.WriteLine("Use: Nadabas.Setup plan-install --source <absolute-xlam-path> --addin-directory <absolute-path>");
            return 2;
        }
        try
        {
            if (!Path.IsPathFullyQualified(args[2]) || !Path.IsPathFullyQualified(args[4]))
                throw new ArgumentException();
            var plan = InstallationPlanner.Create(args[2], args[4]);
            output.WriteLine(JsonSerializer.Serialize(plan, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            }));
            return plan.Status == "blocked" ? 1 : 0;
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            error.WriteLine("Source and add-in directory must be valid absolute paths.");
            return 2;
        }
    }
}
