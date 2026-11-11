using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using NetSparkleUpdater;
using NetSparkleUpdater.Configurations;
using NetSparkleUpdater.Enums;
using NetSparkleUpdater.Interfaces;
using NetSparkleUpdater.SignatureVerifiers;

// Deliberately a local integration probe, not a production updater. No execution,
// scheduling or registry configuration. Trust key is supplied by the test harness.
internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "--version")
        {
            Console.WriteLine(Assembly.GetExecutingAssembly().GetName().Version);
            return 0;
        }
        if (args.Length != 6 || args[0] != "--feed" || args[2] != "--public-key" || args[4] != "--output")
        {
            Console.Error.WriteLine("TEST ONLY: --feed <http://127.0.0.1:port/appcast.xml> --public-key <file> --output <new-absolute-directory>");
            return 2;
        }
        try
        {
            var feed = RequireLoopback(args[1]);
            var key = File.ReadAllText(args[3]).Trim();
            if (Convert.FromBase64String(key).Length != 32) throw new IOException("Invalid test public key.");
            if (!Path.IsPathFullyQualified(args[5]) || Directory.Exists(args[5]) || File.Exists(args[5]))
                throw new IOException("Output must be a new absolute directory.");
            using var updater = new SparkleUpdater(feed.AbsoluteUri, new Ed25519Checker(SecurityMode.Strict, key))
            {
                Configuration = new DefaultConfiguration(new ProbeAssembly()),
                UIFactory = null,
                ShouldKillParentProcessWhenStartingInstaller = false
            };
            var info = await updater.CheckForUpdatesQuietly().WaitAsync(TimeSpan.FromSeconds(30));
            if (info.Status == UpdateStatus.UpdateNotAvailable)
            {
                Report(new { kind = "netsparkle-probe", status = "up-to-date", installedVersion = new ProbeAssembly().AssemblyVersion });
                return 0;
            }
            if (info.Status != UpdateStatus.UpdateAvailable || info.Updates.Count == 0)
                throw new IOException("NetSparkle did not accept an update feed: " + info.Status);
            var update = info.Updates[0];
            var download = RequireLoopback(update.DownloadLink ?? "");
            if (download.Authority != feed.Authority) throw new IOException("Test package must use the feed's loopback origin.");
            using var http = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false }) { Timeout = TimeSpan.FromSeconds(60) };
            // Read to memory, verify with NetSparkle, then write only validated bytes.
            // Production will need bounded streaming and a compiled-in publisher key.
            var bytes = await http.GetByteArrayAsync(download);
            var checker = new Ed25519Checker(SecurityMode.Strict, key);
            if (checker.VerifySignature(update.DownloadSignature ?? "", bytes) != ValidationResult.Valid)
                throw new IOException("Package signature rejected.");
            Directory.CreateDirectory(args[5]);
            var package = Path.Combine(args[5], "verified-prototype-setup.exe");
            using (var file = new FileStream(package, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                await file.WriteAsync(bytes);
            Report(new
            {
                kind = "netsparkle-probe",
                status = "verified-download",
                version = update.Version,
                package,
                sha256 = Convert.ToHexString(SHA256.HashData(bytes)),
                installerExecuted = false
            });
            return 0;
        }
        catch (Exception ex)
        {
            Report(new { kind = "netsparkle-probe", status = "rejected", error = ex.Message, installerExecuted = false });
            return 1;
        }
    }

    private static Uri RequireLoopback(string value)
    {
        var uri = new Uri(value, UriKind.Absolute);
        if (uri.Scheme != "http" || uri.Host != "127.0.0.1" || uri.UserInfo.Length != 0)
            throw new IOException("This prototype accepts only literal 127.0.0.1 HTTP feeds/packages.");
        return uri;
    }

    private static void Report<T>(T value) => Console.WriteLine(JsonSerializer.Serialize(value,
        new JsonSerializerOptions { WriteIndented = true }));

    private sealed class ProbeAssembly : IAssemblyAccessor
    {
        public string AssemblyCompany => "NADABAS development";
        public string AssemblyCopyright => "";
        public string AssemblyDescription => "Isolated installer/update experiment";
        public string AssemblyTitle => "NADABAS Installation Prototype";
        public string AssemblyProduct => AssemblyTitle;
        public string AssemblyVersion => Assembly.GetExecutingAssembly().GetName().Version!.ToString();
    }
}
