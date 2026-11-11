using Chaos.NaCl;
using Nadabas.PackageUpdate;
using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace Nadabas.Setup.Tests;

public sealed class SignedDownloadTransportTests
{
    [Fact]
    public async Task SignedFeedDownloadsThroughRealTlsAndGithubAssetRedirect()
    {
        Ed25519.KeyPairFromSeed(out var publicKey, out var privateKey, RandomNumberGenerator.GetBytes(32));
        byte[] package = [1, 7, 9, 3];
        const string feedUrl = "https://github.com/statisticsnorway/nadabas/releases/latest/download/appcast.xml";
        const string packageUrl = "https://github.com/statisticsnorway/nadabas/releases/download/test/setup.exe";
        string Sign(byte[] bytes) => Convert.ToBase64String(Ed25519.Sign(bytes, privateKey));
        var feed = Encoding.UTF8.GetBytes($"""
            <rss xmlns:sparkle="http://www.andymatuschak.org/xml-namespaces/sparkle"><channel><item>
            <enclosure url="{packageUrl}" sparkle:version="1.0.0" length="{package.Length}" sparkle:signature="{Sign(package)}"/>
            </item></channel></rss>
            """);
        using var rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=github.com", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using var generated = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-1), DateTimeOffset.UtcNow.AddHours(1));
        // Schannel needs a temporary key container. DefaultKeySet releases it
        // with the imported certificate; no certificate trust store is changed.
        var pfx = generated.Export(X509ContentType.Pfx);
        using var certificate = X509CertificateLoader.LoadPkcs12(pfx, null);
        CryptographicOperations.ZeroMemory(pfx);
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var endpoint = (IPEndPoint)listener.LocalEndpoint;
        var paths = new List<string>();
        var server = Serve();
        using var handler = new SocketsHttpHandler
        {
            AllowAutoRedirect = false,
            UseProxy = false,
            ConnectCallback = async (_, cancellation) =>
            {
                var socket = new Socket(SocketType.Stream, ProtocolType.Tcp);
                try
                {
                    await socket.ConnectAsync(endpoint, cancellation);
                    return new NetworkStream(socket, ownsSocket: true);
                }
                catch { socket.Dispose(); throw; }
            },
            // Test transport only. The ephemeral certificate is never installed.
            SslOptions = new SslClientAuthenticationOptions
            {
                RemoteCertificateValidationCallback = (_, peer, _, _) =>
                    peer?.GetCertHashString() == certificate.GetCertHashString()
            }
        };
        using var http = new HttpClient(handler);
        try
        {
            var downloader = new SignedFeed(http, new(feedUrl, Convert.ToBase64String(publicKey)));
            var release = await downloader.Discover("0.9.3", null, deadline.Token);
            Assert.NotNull(release);
            Assert.Equal(package, await downloader.Download(release, deadline.Token));
            await server;
            Assert.Equal(4, paths.Count);
            Assert.EndsWith("/appcast.xml", paths[0]);
            Assert.EndsWith("/appcast.xml.signature", paths[1]);
            Assert.Equal("/fixture.exe", paths[3]);
        }
        finally
        {
            deadline.Cancel();
            listener.Stop();
            try { await server; }
            catch (Exception ex) when (ex is OperationCanceledException or SocketException or IOException) { }
        }

        async Task Serve()
        {
            for (var i = 0; i < 4; i++)
            {
                using var client = await listener.AcceptTcpClientAsync(deadline.Token);
                using var ssl = new SslStream(client.GetStream());
                await ssl.AuthenticateAsServerAsync(new SslServerAuthenticationOptions
                {
                    ServerCertificate = certificate,
                    EnabledSslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13
                }, deadline.Token);
                using var reader = new StreamReader(ssl, Encoding.ASCII, leaveOpen: true);
                var first = await reader.ReadLineAsync(deadline.Token) ?? throw new IOException("Missing HTTP request.");
                paths.Add(first.Split(' ')[1]);
                while (!string.IsNullOrEmpty(await reader.ReadLineAsync(deadline.Token))) { }
                byte[] body = i switch { 0 => feed, 1 => Encoding.UTF8.GetBytes(Sign(feed)), 3 => package, _ => [] };
                var status = i == 2 ? "302 Found" : "200 OK";
                var location = i == 2 ? "Location: https://release-assets.githubusercontent.com/fixture.exe\r\n" : "";
                var header = Encoding.ASCII.GetBytes($"HTTP/1.1 {status}\r\n{location}Content-Length: {body.Length}\r\nConnection: close\r\n\r\n");
                await ssl.WriteAsync(header, deadline.Token);
                await ssl.WriteAsync(body, deadline.Token);
            }
        }
    }
}
