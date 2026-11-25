using System.Net;
using System.Security.Cryptography;
using System.Text;
using Chaos.NaCl;
using Nadabas.PackageUpdate;

namespace Nadabas.Setup.Tests;

public sealed class SignedFeedTests
{
    private const string FeedUrl = "https://github.com/statisticsnorway/nadabas/releases/latest/download/appcast.xml";
    private const string PackageUrl = "https://github.com/statisticsnorway/nadabas/releases/download/v0.9.0/setup.exe";
    private readonly byte[] publicKey;
    private readonly byte[] privateKey;
    public SignedFeedTests() => Ed25519.KeyPairFromSeed(out publicKey, out privateKey, RandomNumberGenerator.GetBytes(32));
    private string Sign(byte[] bytes) => Convert.ToBase64String(Ed25519.Sign(bytes, privateKey));
    private byte[] Feed(byte[] package, string version = "0.9.0", string? url = null) => Encoding.UTF8.GetBytes($"""
        <rss xmlns:sparkle="http://www.andymatuschak.org/xml-namespaces/sparkle"><channel><item>
        <enclosure url="{url ?? PackageUrl}" sparkle:version="{version}" length="{package.Length}" sparkle:os="windows" sparkle:signature="{Sign(package)}" />
        </item></channel></rss>
        """);

    private sealed class Handler(Func<HttpRequestMessage, HttpResponseMessage> respond) : HttpMessageHandler
    {
        internal int Calls;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            Calls++;
            return Task.FromResult(respond(request));
        }
    }
    private static HttpResponseMessage Bytes(byte[] bytes) => new(HttpStatusCode.OK) { Content = new ByteArrayContent(bytes) };
    private SignedFeed Client(HttpClient http) => new(http, new(FeedUrl, Convert.ToBase64String(publicKey)));

    [Fact]
    public async Task AuthenticatedFeedAndPackageDownloadSucceed()
    {
        byte[] package = [1, 2, 3];
        var feed = Feed(package);
        using var http = new HttpClient(new Handler(r => r.RequestUri!.AbsoluteUri switch
        {
            FeedUrl => Bytes(feed),
            FeedUrl + ".signature" => Bytes(Encoding.UTF8.GetBytes(Sign(feed))),
            PackageUrl => Bytes(package),
            _ => throw new InvalidOperationException()
        }));
        var client = Client(http);
        var release = await client.Discover("0.8.2", null, default);
        Assert.NotNull(release);
        Assert.Equal(package, await client.Download(release, default));
    }

    [Fact]
    public async Task ModifiedFeedRejectedBeforeAnyPackageRequest()
    {
        var feed = Feed([1]);
        var signature = Sign(feed);
        var handler = new Handler(r => Bytes(r.RequestUri!.AbsoluteUri.EndsWith(".signature") ? Encoding.UTF8.GetBytes(signature) : feed.Concat(new byte[] { 32 }).ToArray()));
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<IOException>(() => Client(http).Discover("0.8.2", null, default));
        Assert.Equal(2, handler.Calls);
    }

    [Fact]
    public async Task WrongKeyAndChangedPackageRejected()
    {
        byte[] original = [1, 2];
        using var http = new HttpClient(new Handler(_ => Bytes([1, 3])));
        await Assert.ThrowsAsync<IOException>(() => Client(http).Download(new("0.9.0", new(PackageUrl), 2, Sign(original)), default));
        Ed25519.KeyPairFromSeed(out var otherPublic, out _, RandomNumberGenerator.GetBytes(32));
        var wrongKey = new SignedFeed(http, new(FeedUrl, Convert.ToBase64String(otherPublic)));
        await Assert.ThrowsAsync<IOException>(() => wrongKey.Download(new("0.9.0", new(PackageUrl), 2, Sign([1, 3])), default));
    }

    [Theory]
    [InlineData("http://github.com/statisticsnorway/nadabas/releases/download/v1/setup.exe")]
    [InlineData("https://github.com/other/repo/releases/download/v1/setup.exe")]
    [InlineData("https://github.com.evil.test/statisticsnorway/nadabas/releases/download/v1/setup.exe")]
    [InlineData("https://user:password@github.com/statisticsnorway/nadabas/releases/download/v1/setup.exe")]
    [InlineData("https://github.com:444/statisticsnorway/nadabas/releases/download/v1/setup.exe")]
    public void UnapprovedUrlsAreRejected(string url) => Assert.Throws<IOException>(() => SignedFeed.RequireUrl(new(url)));

    [Fact]
    public async Task RedirectCannotEscapeAllowedHosts()
    {
        var handler = new Handler(_ => { var r = new HttpResponseMessage(HttpStatusCode.Redirect); r.Headers.Location = new("https://evil.test/setup.exe"); return r; });
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<IOException>(() => Client(http).ReadBounded(new(PackageUrl), 100, default));
        Assert.Equal(1, handler.Calls);
    }

    [Fact]
    public async Task GithubReleaseAssetRedirectAccepted()
    {
        using var http = new HttpClient(new Handler(r =>
        {
            if (r.RequestUri!.Host == "release-assets.githubusercontent.com") return Bytes([42]);
            var response = new HttpResponseMessage(HttpStatusCode.Redirect);
            response.Headers.Location = new("https://release-assets.githubusercontent.com/asset?token=example");
            return response;
        }));
        Assert.Equal(new byte[] { 42 }, await Client(http).ReadBounded(new(PackageUrl), 1, default));
    }

    [Fact]
    public async Task ExcessLengthAndCancellationStopDownload()
    {
        using var http = new HttpClient(new Handler(_ => Bytes([1, 2, 3])));
        await Assert.ThrowsAsync<IOException>(() => Client(http).ReadBounded(new(PackageUrl), 2, default));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Client(http).ReadBounded(new(PackageUrl), 100, cancelled.Token));
    }

    [Fact]
    public void SameOlderAndNonPinnedVersionsAreNotOffered()
    {
        var feed = Feed([1]);
        Assert.Null(SignedFeed.Select(feed, "0.9.0", null));
        Assert.Null(SignedFeed.Select(feed, "1.0.0", null));
        Assert.Null(SignedFeed.Select(feed, "0.8.2", "0.8.3"));
        Assert.NotNull(SignedFeed.Select(feed, "0.8.2", "0.9.0.0"));
    }

    [Fact]
    public void XmlEntityExpansionIsProhibited()
    {
        Assert.Throws<System.Xml.XmlException>(() => SignedFeed.Select(Encoding.UTF8.GetBytes("<!DOCTYPE rss [<!ENTITY x SYSTEM 'file:///secret'>]><rss>&x;</rss>"), "0.8.2", null));
    }

    [Fact]
    public void PolicyCanBlockManualAndAutomaticUpdatesIndependently()
    {
        Assert.Throws<IOException>(() => new Nadabas.PackageUpdate.UpdatePolicy(false, true, null).RequireAllowed(true, "0.9.0"));
        Assert.Throws<IOException>(() => new Nadabas.PackageUpdate.UpdatePolicy(true, false, null).RequireAllowed(false));
        new Nadabas.PackageUpdate.UpdatePolicy(true, false, null).RequireAllowed(true);
        Assert.Throws<IOException>(() => new Nadabas.PackageUpdate.UpdatePolicy(true, true, "0.8.2").RequireAllowed(true, "0.9.0"));
    }
}
