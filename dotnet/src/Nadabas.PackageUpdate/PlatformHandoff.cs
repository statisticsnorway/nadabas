using System.Diagnostics;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text.Json;

namespace Nadabas.PackageUpdate;

public sealed record PlatformOffer(string Id, string Version);

// A pointer selects a job. Only the verified signed files in that job can
// authorize execution. All paths below are derived from the installed root.
[SupportedOSPlatform("windows")]
public static class PlatformHandoff
{
    private const string Pointer = "pending-platform-update";

    private static string PathIn(string directory, string name) => PackageRules.SafePath(Path.Combine(directory, name));
    internal static string Cache(string root, string id)
    {
        if (!Guid.TryParseExact(id, "N", out _)) throw new IOException("Invalid package job identity.");
        return PathIn(PackageRules.SafePath(root) + "-package-cache", id);
    }

    public static string? ReadId(string root)
    {
        var file = PathIn(root, Pointer);
        if (!File.Exists(file)) return null;
        if (new FileInfo(file).Length != 32) throw new IOException("Invalid package pointer.");
        var id = File.ReadAllText(file);
        _ = Cache(root, id);
        return id;
    }

    private static FileStream Lock(string root) => new(PathIn(root, "platform-update.lock"),
        FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);

    internal static PlatformOffer Stage(string root, string installer, bool smoke)
    {
        root = PackageRules.SafePath(root);
        _ = PackageJob.InstalledVersion(root, smoke);
        using var gate = Lock(root);
        if (ReadId(root) is not null) throw new IOException("An existing package job needs attention first.");
        var id = Guid.NewGuid().ToString("N");
        var job = PackageJob.Stage(root, installer, Cache(root, id), smoke);
        var temporary = PathIn(root, id + ".pending");
        File.WriteAllText(temporary, id);
        File.Move(temporary, PathIn(root, Pointer));
        return new(id, job.TargetVersion);
    }

    private static PendingPackage ReadJob(string root, string id, bool requirePending = true)
    {
        if (ReadId(root) != id) throw new IOException("The pending package has changed.");
        var cache = Cache(root, id);
        var path = PathIn(cache, "package-update.json");
        if (new FileInfo(path).Length > 8192) throw new IOException("Invalid package job size.");
        var job = JsonSerializer.Deserialize<PendingPackage>(File.ReadAllText(path));
        if (job is null || job.SchemaVersion != 1 ||
            !string.Equals(PackageRules.SafePath(job.Root), PackageRules.SafePath(root), StringComparison.OrdinalIgnoreCase))
            throw new IOException("Package job does not belong to this installation.");
        if (requirePending && File.Exists(PathIn(cache, "apply-started.json"))) throw new IOException("This package job already started. Inspect the installer log.");
        if (File.Exists(PathIn(cache, "cancelled.json"))) throw new IOException("This package job was cancelled.");
        return job;
    }

    public static PlatformOffer Verify(string root, string id)
    {
        var job = ReadJob(root, id);
        if (!job.Smoke) UpdatePolicy.Read().RequireAllowed(true, job.TargetVersion);
        var current = PackageJob.InstalledVersion(root, job.Smoke);
        if (PackageRules.Normalize(current) != PackageRules.Normalize(job.PreviousVersion))
            throw new IOException("The installed platform has changed.");
        PackageRules.RequireUpgrade(current, job.TargetVersion);
        var file = PathIn(Cache(root, id), "setup.exe");
        using var lease = File.OpenRead(file);
        if (Convert.ToHexString(SHA256.HashData(lease)) != job.Sha256 ||
            PackageJob.VerifyPackage(file, job.Smoke) != job.TargetVersion)
            throw new IOException("The cached installer failed verification.");
        return new(id, job.TargetVersion);
    }

    public static void Cancel(string root, string id)
    {
        using var gate = Lock(root);
        if (ReadId(root) != id) throw new IOException("The pending package has changed.");
        var cache = Cache(root, id);
        using var jobGate = new FileStream(PathIn(cache, "apply.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        if (File.Exists(PathIn(cache, "apply-started.json"))) throw new IOException("Package installation has already started.");
        var pending = ReadJob(root, id);
        File.WriteAllText(PathIn(root, "skipped-package-version"), pending.TargetVersion);
        File.WriteAllText(PathIn(cache, "cancelled.json"), "{}");
        File.Delete(PathIn(root, Pointer));
    }

    internal static async Task<ApplyResult> Recover(string root, string id, string previousInstaller)
    {
        root = PackageRules.SafePath(root);
        var cache = Cache(root, id);
        PackageRules.RequireOutside(root, Environment.ProcessPath!);
        using var gate = Lock(root);
        var job = ReadJob(root, id, requirePending: false);
        using var jobGate = new FileStream(PathIn(cache, "apply.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        if (!File.Exists(PathIn(cache, "apply-started.json"))) throw new IOException("Recovery requires a previously started installation.");
        if (PackageJob.AnyExcel()) return new("deferred-excel-running");
        var current = PackageRules.Normalize(PackageJob.InstalledVersion(root, job.Smoke));
        if (current != PackageRules.Normalize(job.TargetVersion) && current != PackageRules.Normalize(job.PreviousVersion))
            throw new IOException("Another platform version was installed. Review it before recovery.");
        previousInstaller = PackageRules.SafePath(previousInstaller);
        PackageRules.RequireOutside(root, previousInstaller);
        using var lease = File.OpenRead(previousInstaller);
        if (PackageRules.Normalize(PackageJob.VerifyPackage(previousInstaller, job.Smoke)) != PackageRules.Normalize(job.PreviousVersion))
            throw new IOException("Recovery requires the verified original platform version.");
        if (PackageJob.AnyExcel()) return new("deferred-excel-running");
        try
        {
            var version = await PackageJob.RunSetup(root, job.Smoke, previousInstaller, job.PreviousVersion, PathIn(cache, "recovery.log"));
            PackageJob.Write(cache, "status.json", new { status = "restored", version });
            File.Delete(PathIn(root, Pointer));
            return new("restored", version);
        }
        catch (Exception ex)
        {
            PackageJob.Write(cache, "status.json", new { status = "repair-required", error = ex.Message });
            throw;
        }
    }

    public static Process LaunchWaiter(string root, string id, Process excel)
    {
        _ = Verify(root, id);
        var path = PathIn(Cache(root, id), "Nadabas.PackageUpdate.exe");
        using var lease = File.OpenRead(path);
        using var installed = File.OpenRead(PathIn(root, "updater/Nadabas.PackageUpdate.exe"));
        if (!SHA256.HashData(lease).SequenceEqual(SHA256.HashData(installed)))
            throw new IOException("The cached helper differs from the installed helper.");
        PublisherTrust.Verify(path);
        var start = new ProcessStartInfo(path) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (var argument in new[] { "wait-apply", root, id, excel.Id.ToString(System.Globalization.CultureInfo.InvariantCulture),
            excel.StartTime.ToUniversalTime().Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture) }) start.ArgumentList.Add(argument);
        return Process.Start(start) ?? throw new IOException("Could not start the external update helper.");
    }

    internal static async Task WaitAndApply(string root, string id, int pid, long ticks)
    {
        var cache = Cache(root, id);
        PackageRules.RequireCachedExecutable(Environment.ProcessPath ?? throw new IOException("Cannot identify the restart helper."),
            PathIn(cache, "Nadabas.PackageUpdate.exe"));
        using var parent = Process.GetProcessById(pid);
        var executable = parent.MainModule?.FileName;
        if (parent.StartTime.ToUniversalTime().Ticks != ticks ||
            !string.Equals(Path.GetFileName(executable), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
            throw new IOException("The original Excel process could not be identified.");
        _ = parent.SafeHandle;
        // The launcher already verified the job. Check its identity again before
        // acknowledging. Apply performs full signature verification after exit.
        _ = ReadJob(root, id);
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(60));
        Console.WriteLine("restart-helper-ready");
        try
        {
            await PackageRestart.AfterExit(parent.WaitForExitAsync, PackageJob.AnyExcel, async () =>
            {
                using var gate = Lock(root);
                _ = ReadJob(root, id);
                var result = await PackageJob.Apply(cache);
                if (result.Status != "installed") return false;
                File.Delete(PathIn(root, Pointer));
                return true;
            }, () => Process.Start(new ProcessStartInfo(executable!) { UseShellExecute = true, Arguments = "/x" }), deadline.Token);
        }
        catch (OperationCanceledException)
        {
            PackageJob.Write(cache, "restart-status.json", new { status = "excel-did-not-exit", retryAllowed = true });
        }
        catch (Exception ex)
        {
            PackageJob.Write(cache, "restart-status.json", new { status = "restart-failed", error = ex.Message });
            throw;
        }
    }
}
