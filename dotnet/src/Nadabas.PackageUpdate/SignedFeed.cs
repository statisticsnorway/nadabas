using System.Net;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using NetSparkleUpdater.Enums;
using NetSparkleUpdater.SignatureVerifiers;

namespace Nadabas.PackageUpdate;

internal sealed record UpdateChannel(string Feed, string PublicKey,
    string[]? AdditionalPublicKeys = null, string[]? RevokedPublicKeys = null)
{
    internal string[] TrustedKeys()
    {
        var configured = new[] { PublicKey }.Concat(AdditionalPublicKeys ?? []).ToArray();
        var revoked = RevokedPublicKeys ?? [];
        foreach (var key in configured.Concat(revoked))
        {
            var bytes = Convert.FromBase64String(key);
            if (bytes.Length != 32 || Convert.ToBase64String(bytes) != key)
                throw new IOException("Invalid update verification key.");
        }
        var trusted = configured.Except(revoked, StringComparer.Ordinal).Distinct(StringComparer.Ordinal).ToArray();
        if (trusted.Length == 0) throw new IOException("No trusted update verification keys.");
        return trusted;
    }

    internal bool Verify(string signature, byte[] content) => TrustedKeys().Any(key =>
        new Ed25519Checker(SecurityMode.Strict, key).VerifySignature(signature, content) == ValidationResult.Valid);

    internal static UpdateChannel Bundled()
    {
        using var stream = typeof(UpdateChannel).Assembly.GetManifestResourceStream("Nadabas.UpdateChannel.json")!;
        return System.Text.Json.JsonSerializer.Deserialize<UpdateChannel>(stream)!;
    }
}

internal sealed record FeedRelease(string Version, Uri Url, long Length, string Signature);

// NetSparkle's appcast/signature format, with explicit network and size bounds.
// Trust comes only from the signed updater assembly, never a downloaded sidecar.
internal sealed class SignedFeed(HttpClient http, UpdateChannel channel)
{
    internal const int MaximumPackageBytes = 128 * 1024 * 1024;
    internal const int MaximumFeedBytes = 1024 * 1024;
    private const string Repository = "/statisticsnorway/nadabas/releases/";
    private static readonly XNamespace Sparkle = "http://www.andymatuschak.org/xml-namespaces/sparkle";

    internal static void RequireUrl(Uri url, bool redirect = false)
    {
        if (url.Scheme != "https" || !url.IsDefaultPort || url.UserInfo.Length != 0 || url.Fragment.Length != 0)
            throw new IOException("Updates require an approved HTTPS URL.");
        if (url.Host == "github.com" && url.AbsolutePath.StartsWith(Repository, StringComparison.Ordinal) &&
            (url.AbsolutePath.StartsWith(Repository + "download/", StringComparison.Ordinal) ||
             url.AbsolutePath.StartsWith(Repository + "latest/download/", StringComparison.Ordinal))) return;
        if (redirect && url.Host == "release-assets.githubusercontent.com") return;
        throw new IOException("Update URL is outside the approved release repository.");
    }

    internal async Task<byte[]> ReadBounded(Uri url, int limit, CancellationToken cancellation)
    {
        RequireUrl(url);
        for (var redirect = 0; redirect <= 5; redirect++)
        {
            RequireUrl(url, redirect > 0);
            using var response = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellation).ConfigureAwait(false);
            if (response.StatusCode is HttpStatusCode.Moved or HttpStatusCode.Redirect or HttpStatusCode.RedirectMethod or HttpStatusCode.TemporaryRedirect or HttpStatusCode.PermanentRedirect)
            {
                url = new Uri(url, response.Headers.Location ?? throw new IOException("Redirect lacks a destination."));
                continue;
            }
            response.EnsureSuccessStatusCode();
            if (response.Content.Headers.ContentLength > limit) throw new IOException("Update response is too large.");
            await using var input = await response.Content.ReadAsStreamAsync(cancellation).ConfigureAwait(false);
            using var output = new MemoryStream();
            var buffer = new byte[65536];
            int count;
            while ((count = await input.ReadAsync(buffer, cancellation).ConfigureAwait(false)) != 0)
            {
                if (output.Length + count > limit) throw new IOException("Update response exceeded its size limit.");
                output.Write(buffer, 0, count);
            }
            return output.ToArray();
        }
        throw new IOException("Too many update redirects.");
    }

    internal async Task<FeedRelease?> Discover(string current, string? pin, CancellationToken cancellation)
    {
        _ = channel.TrustedKeys();
        var uri = new Uri(channel.Feed);
        var feed = await ReadBounded(uri, MaximumFeedBytes, cancellation).ConfigureAwait(false);
        var signature = Encoding.UTF8.GetString(await ReadBounded(new Uri(channel.Feed + ".signature"), 1024, cancellation).ConfigureAwait(false)).Trim();
        if (!channel.Verify(signature, feed)) throw new IOException("Feed signature rejected.");
        return Select(feed, current, pin);
    }

    internal static FeedRelease? Select(byte[] verifiedFeed, string current, string? pin)
    {
        using var input = new MemoryStream(verifiedFeed);
        using var reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = MaximumFeedBytes });
        var document = XDocument.Load(reader);
        if (document.Root?.Name != "rss") throw new IOException("Invalid appcast.");
        var releases = new List<FeedRelease>();
        foreach (var item in document.Root.Elements("channel").Elements("item"))
        {
            if (item.Element(Sparkle + "channel") is { Value: not "stable" }) continue;
            foreach (var enclosure in item.Elements("enclosure"))
            {
                if ((string?)enclosure.Attribute(Sparkle + "os") is not (null or "windows" or "win")) continue;
                var version = (string?)enclosure.Attribute(Sparkle + "version") ?? throw new IOException("Missing package version.");
                var parsed = PackageRules.Normalize(version);
                if (parsed <= PackageRules.Normalize(current) || (pin is not null && parsed != PackageRules.Normalize(pin))) continue;
                var url = new Uri((string?)enclosure.Attribute("url") ?? "", UriKind.Absolute);
                RequireUrl(url);
                if (!long.TryParse((string?)enclosure.Attribute("length"), out var length) || length <= 0 || length > MaximumPackageBytes)
                    throw new IOException("Invalid package size.");
                var signature = (string?)enclosure.Attribute(Sparkle + "signature") ?? (string?)enclosure.Attribute(Sparkle + "edSignature") ?? "";
                if (Convert.FromBase64String(signature).Length != 64) throw new IOException("Invalid package signature.");
                releases.Add(new(version, url, length, signature));
            }
        }
        var selected = releases.OrderByDescending(x => PackageRules.Normalize(x.Version)).FirstOrDefault();
        if (selected is not null && releases.Count(x => PackageRules.Normalize(x.Version) == PackageRules.Normalize(selected.Version)) != 1)
            throw new IOException("Ambiguous release version.");
        return selected;
    }

    internal async Task<byte[]> Download(FeedRelease release, CancellationToken cancellation)
    {
        var bytes = await ReadBounded(release.Url, checked((int)release.Length), cancellation).ConfigureAwait(false);
        if (bytes.LongLength != release.Length || !channel.Verify(release.Signature, bytes))
            throw new IOException("Downloaded package signature or length rejected.");
        return bytes;
    }
}
