using System.Diagnostics;
using System.Runtime.Versioning;
using System.Text.Json;

namespace Nadabas.PackageUpdate;

public sealed record DownloadStatus(string Status, string? Version = null, UpdateDiagnostic? Diagnostic = null);

[SupportedOSPlatform("windows")]
public static class AutomaticUpdates
{
    public static async Task<DownloadStatus> CheckFromExcel(string root, bool manual, CancellationToken cancellation)
    {
        var path = PackageRules.SafePath(Path.Combine(root, "updater", "Nadabas.PackageUpdate.exe"));
        using var lease = File.OpenRead(path);
        PublisherTrust.Verify(path);
        var command = new ProcessStartInfo(path) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        command.ArgumentList.Add("check-download");
        command.ArgumentList.Add(root);
        if (manual) command.ArgumentList.Add("--manual");
        using var process = Process.Start(command) ?? throw new IOException("Cannot start update checker.");
        try
        {
            var output = process.StandardOutput.ReadToEndAsync(cancellation);
            var error = process.StandardError.ReadToEndAsync(cancellation);
            await process.WaitForExitAsync(cancellation).ConfigureAwait(false);
            var result = await output.ConfigureAwait(false);
            _ = await error.ConfigureAwait(false);
            var status = JsonSerializer.Deserialize<DownloadStatus>(result) ?? throw new IOException("Missing update result.");
            if (process.ExitCode != 0 && !(process.ExitCode == 1 && status.Status == "failed" && status.Diagnostic is not null))
                throw new IOException("Update check failed. The installed version remains unchanged.");
            return status;
        }
        catch
        {
            if (!process.HasExited) process.Kill(); // Only the checker; it never executes Setup.
            throw;
        }
    }

    internal static async Task<DownloadStatus> Check(string root, bool manual)
    {
        root = PackageRules.SafePath(root);
        var current = PackageJob.InstalledVersion(root, false);
        var policy = UpdatePolicy.Read();
        if (!policy.Enabled || (!manual && !policy.Automatic)) return new("disabled");
        var channel = UpdateChannel.Bundled();
        if (string.IsNullOrWhiteSpace(channel.PublicKey)) return new("not-configured");
        // Serialize across Excel processes; an existing checker may finish normally.
        FileStream gate;
        try { gate = new FileStream(PackageRules.SafePath(Path.Combine(root, "download.lock")), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
        catch (IOException) { return new("checking"); }
        using (gate)
        {
            if (PlatformHandoff.ReadId(root) is not null) return new("ready");
            var attemptPath = PackageRules.SafePath(Path.Combine(root, "last-download-check.json"));
            if (!manual && File.Exists(attemptPath) && new FileInfo(attemptPath).Length < 2048)
            {
                var last = JsonSerializer.Deserialize<DateTimeOffset>(File.ReadAllText(attemptPath));
                if (last <= DateTimeOffset.UtcNow && DateTimeOffset.UtcNow - last < TimeSpan.FromHours(6)) return new("deferred");
            }
            PackageJob.Write(root, "last-download-check.json", DateTimeOffset.UtcNow);
            using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3));
            using var http = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false }) { Timeout = TimeSpan.FromSeconds(60) };
            http.DefaultRequestHeaders.UserAgent.ParseAdd("NADABAS-Updater/0.9");
            var feed = new SignedFeed(http, channel);
            var release = await feed.Discover(current, policy.PinnedVersion, deadline.Token).ConfigureAwait(false);
            if (release is null) return new("up-to-date");
            var skippedPath = PackageRules.SafePath(Path.Combine(root, "skipped-package-version"));
            if (!manual && File.Exists(skippedPath) && new FileInfo(skippedPath).Length < 128 &&
                PackageRules.Normalize(File.ReadAllText(skippedPath)) == PackageRules.Normalize(release.Version)) return new("skipped", release.Version);
            UpdatePolicy.Read().RequireAllowed(manual, release.Version);
            var bytes = await feed.Download(release, deadline.Token).ConfigureAwait(false);
            var directory = PackageRules.SafePath(root + "-downloads");
            Directory.CreateDirectory(directory);
            var path = PackageRules.SafePath(Path.Combine(directory, Guid.NewGuid().ToString("N") + ".exe"));
            try
            {
                await using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                    await file.WriteAsync(bytes, deadline.Token).ConfigureAwait(false);
                using var lease = File.OpenRead(path);
                if (PackageRules.Normalize(PackageJob.VerifyPackage(path, false)) != PackageRules.Normalize(release.Version))
                    throw new IOException("Signed installer version does not match the signed feed.");
                UpdatePolicy.Read().RequireAllowed(manual, release.Version);
                _ = PlatformHandoff.Stage(root, path, false);
                return new("ready", release.Version);
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }
    }
}
