using Chaos.NaCl;
using Nadabas.PackageUpdate;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

// Release workstation tool only. Never included in the installation payload.
if (!OperatingSystem.IsWindows()) throw new IOException("Windows user protection is required.");
if (args.Length == 1 && args[0] == "check-channel")
{
    // Read-only release preflight, using the same pinned trust and HTTP rules as Excel.
    using var handler = new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false };
    using var http = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(60) };
    using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3));
    try
    {
        var feed = new SignedFeed(http, UpdateChannel.Bundled());
        var release = await feed.Discover("0.0.0", null, deadline.Token);
        Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(new
        {
            status = release is null ? "no-supported-release" : "verified-feed",
            version = release?.Version,
            packageDownloaded = false,
            installationChanged = false
        }));
    }
    catch (Exception error)
    {
        Console.WriteLine(UpdateDiagnostic.Create(UpdateStage.Download, error).ToJson());
        Environment.ExitCode = 1;
    }
    return;
}
if ((args.Length == 3 && args[0] == "backup-key") || (args.Length == 4 && args[0] == "restore-key"))
{
    Nadabas.FeedSigning.KeyRecovery.Run(args[0], args[1], args[2], args.ElementAtOrDefault(3));
    return;
}
if (args.Length == 2 && args[0] == "verify-feed")
{
    var directory = PackageRules.SafePath(args[1]);
    var channel = UpdateChannel.Bundled();
    var feedPath = PackageRules.SafePath(Path.Combine(directory, "appcast.xml"));
    if (new FileInfo(feedPath).Length > SignedFeed.MaximumFeedBytes) throw new IOException("Feed too large.");
    var feed = File.ReadAllBytes(feedPath);
    var signature = File.ReadAllText(PackageRules.SafePath(feedPath + ".signature")).Trim();
    if (!channel.Verify(signature, feed)) throw new IOException("Invalid feed signature.");
    var release = SignedFeed.Select(feed, "0.0.0", null) ?? throw new IOException("No supported release.");
    var name = Uri.UnescapeDataString(Path.GetFileName(release.Url.AbsolutePath));
    if (name != Path.GetFileName(name)) throw new IOException("Invalid asset filename.");
    var package = PackageRules.SafePath(Path.Combine(directory, name));
    using var lease = File.OpenRead(package);
    if (lease.Length > SignedFeed.MaximumPackageBytes || lease.Length != release.Length || !channel.Verify(release.Signature, File.ReadAllBytes(package)) ||
        PackageRules.Normalize(PackageJob.VerifyPackage(package, false)) != PackageRules.Normalize(release.Version))
        throw new IOException("Package verification failed.");
    Console.WriteLine("Verified feed signature, package signature, SSB publisher and matching version " + release.Version);
    return;
}
if (args.Length == 1 && args[0] == "new-key")
{
    var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "NADABAS-ReleaseSigning", Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(directory);
    var seed = RandomNumberGenerator.GetBytes(32);
    byte[]? privateKey = null;
    try
    {
        Ed25519.KeyPairFromSeed(out var publicKey, out privateKey, seed);
        var encrypted = ProtectedData.Protect(seed, null, DataProtectionScope.CurrentUser);
        File.WriteAllBytes(Path.Combine(directory, "feed-key.dpapi"), encrypted);
        File.WriteAllText(Path.Combine(directory, "feed-public-key.txt"), Convert.ToBase64String(publicKey));
        Console.WriteLine(directory); // Public location only; never print private material.
    }
    finally
    {
        CryptographicOperations.ZeroMemory(seed);
        if (privateKey is not null) CryptographicOperations.ZeroMemory(privateKey);
    }
    return;
}
if (args.Length == 5 && args[0] == "prepare-feed")
{
    var keyPath = PackageRules.SafePath(args[1]);
    var installer = PackageRules.SafePath(args[2]);
    var tag = args[3];
    if (tag.Length == 0 || tag.Any(c => !(char.IsAsciiLetterOrDigit(c) || c is '.' or '-' or '_')))
        throw new IOException("Use a simple GitHub release tag.");
    var output = PackageRules.SafePath(args[4]);
    if (Directory.Exists(output) || File.Exists(output)) throw new IOException("Output must be new.");
    using var lease = File.OpenRead(installer);
    var version = PackageJob.VerifyPackage(installer, false);
    if (lease.Length > SignedFeed.MaximumPackageBytes) throw new IOException("Package exceeds download limit.");
    var package = File.ReadAllBytes(installer);
    var seed = ProtectedData.Unprotect(File.ReadAllBytes(keyPath), null, DataProtectionScope.CurrentUser);
    byte[]? privateKey = null;
    try
    {
        if (seed.Length != 32) throw new IOException("Invalid key seed.");
        Ed25519.KeyPairFromSeed(out var publicKey, out privateKey, seed);
        var key = Convert.ToBase64String(publicKey);
        if (!UpdateChannel.Bundled().TrustedKeys().Contains(key)) throw new IOException("This key is not trusted in the built updater. Rebuild before preparing a release.");
        XNamespace sparkle = "http://www.andymatuschak.org/xml-namespaces/sparkle";
        var name = Path.GetFileName(installer);
        var url = "https://github.com/statisticsnorway/nadabas/releases/download/" + tag + "/" + Uri.EscapeDataString(name);
        var document = new XDocument(new XElement("rss", new XAttribute("version", "2.0"), new XAttribute(XNamespace.Xmlns + "sparkle", sparkle),
            new XElement("channel", new XElement("title", "NADABAS updates"),
                new XElement("item", new XElement("title", "NADABAS " + version), new XElement(sparkle + "channel", "stable"),
                    new XElement("enclosure", new XAttribute("url", url), new XAttribute(sparkle + "version", version),
                        new XAttribute("length", package.Length), new XAttribute(sparkle + "os", "windows"),
                        new XAttribute(sparkle + "signature", Convert.ToBase64String(Ed25519.Sign(package, privateKey))))))));
        var feed = Encoding.UTF8.GetBytes(document.ToString());
        Directory.CreateDirectory(output);
        File.WriteAllBytes(Path.Combine(output, name), package);
        File.WriteAllBytes(Path.Combine(output, "appcast.xml"), feed);
        File.WriteAllText(Path.Combine(output, "appcast.xml.signature"), Convert.ToBase64String(Ed25519.Sign(feed, privateKey)));
        File.WriteAllText(Path.Combine(output, "release-review.txt"), $"Prepared only; nothing published.\nVersion: {version}\nTag: {tag}\nSHA256: {Convert.ToHexString(SHA256.HashData(package))}\n");
        Console.WriteLine("Prepared signed feed and package for review. Nothing was uploaded.");
    }
    finally
    {
        CryptographicOperations.ZeroMemory(seed);
        if (privateKey is not null) CryptographicOperations.ZeroMemory(privateKey);
    }
    return;
}
Console.Error.WriteLine("check-channel | new-key | verify-feed <directory> | backup-key <feed-key.dpapi> <new-backup-file> | restore-key <backup-file> <new-feed-key.dpapi> <expected-public-key> | prepare-feed <feed-key.dpapi> <SSB-signed-installer> <release-tag> <new-output-directory>");
Environment.ExitCode = 2;
