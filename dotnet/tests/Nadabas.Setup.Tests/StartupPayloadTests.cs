using System.Security.Cryptography;
using Nadabas.ExcelStartup;

namespace Nadabas.Setup.Tests;

[CollectionDefinition("Startup environment", DisableParallelization = true)]
public sealed class StartupEnvironmentCollection;

[Collection("Startup environment")]
public sealed class StartupPayloadTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), "nadabas-startup-" + Guid.NewGuid());

    private string[] Fixture()
    {
        var files = new[] { "Nadabas.Loader.x64.xll", "runtime/x64/dotnet.exe",
            "runtime/x64/host/fxr/10.0.11/hostfxr.dll",
            "runtime/x64/shared/Microsoft.NETCore.App/10.0.11/coreclr.dll",
            "runtime/x64/shared/Microsoft.WindowsDesktop.App/10.0.11/System.Windows.Forms.dll" };
        foreach (var file in files)
        {
            var path = Path.Combine(root, file);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, "isolated fixture: " + file);
        }
        return files.Select(file => "x64\t" + file + "\t" +
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Path.Combine(root, file))))).ToArray();
    }

    [Theory]
    [InlineData(4, "x86")]
    [InlineData(8, "x64")]
    public void SelectsByExcelProcessSize(int size, string expected) =>
        Assert.Equal(expected, StartupPayload.Architecture(size));

    [Fact]
    public void UnconfiguredBuildCannotActivate() =>
        Assert.Throws<IOException>(() => StartupPayload.Verify(root, "x64", ["# placeholder"], _ => Assert.Fail("Publisher must not be reached")));

    [Fact]
    public void OtherArchitectureCannotSupplyMissingManifest()
    {
        var manifest = Fixture();
        Assert.Throws<IOException>(() => StartupPayload.Verify(root, "x86", manifest, _ => Assert.Fail("Wrong architecture")));
    }

    [Theory]
    [InlineData("x86")]
    [InlineData("x64")]
    public void MatchingPayloadIsVerifiedForEitherArchitecture(string architecture)
    {
        var manifest = Fixture();
        if (architecture == "x86")
        {
            File.Move(Path.Combine(root, "Nadabas.Loader.x64.xll"), Path.Combine(root, "Nadabas.Loader.x86.xll"));
            Directory.Move(Path.Combine(root, "runtime/x64"), Path.Combine(root, "runtime/x86"));
            // File bytes remain identical; only the manifest paths and architecture change.
            manifest = manifest.Select(line => line.Replace("x64", "x86", StringComparison.Ordinal)).ToArray();
        }
        var verified = false;
        using var payload = StartupPayload.Verify(root, architecture, manifest, _ => verified = true);
        Assert.True(verified);
        Assert.Equal(Path.Combine(root, "runtime", architecture), payload.Runtime);
    }

    [Fact]
    public void TamperedRuntimeIsRejectedBeforePublisherOrRegistration()
    {
        var manifest = Fixture();
        File.AppendAllText(Path.Combine(root, "runtime/x64/shared/Microsoft.NETCore.App/10.0.11/coreclr.dll"), "tampered");
        Assert.Throws<IOException>(() => StartupPayload.Verify(root, "x64", manifest, _ => Assert.Fail("Unverified payload")));
    }

    [Fact]
    public void UnexpectedRuntimeFileIsRejected()
    {
        var manifest = Fixture();
        File.WriteAllText(Path.Combine(root, "runtime/x64/extra.dll"), "extra");
        Assert.Throws<IOException>(() => StartupPayload.Verify(root, "x64", manifest, _ => Assert.Fail("Unexpected file")));
    }

    [Fact]
    public void IncompleteManifestCannotFallBackToMachineRuntime()
    {
        var manifest = Fixture().Where(line => !line.Contains("hostfxr.dll", StringComparison.Ordinal));
        Assert.Throws<IOException>(() => StartupPayload.Verify(root, "x64", manifest, _ => Assert.Fail("Incomplete runtime")));
    }

    [Fact]
    public void PublisherFailureReleasesFileLeases()
    {
        var manifest = Fixture();
        Assert.Throws<IOException>(() => StartupPayload.Verify(root, "x64", manifest, _ => throw new IOException("Bad publisher")));
        using var replacement = File.Open(Path.Combine(root, "Nadabas.Loader.x64.xll"), FileMode.Open, FileAccess.Write, FileShare.None);
    }

    [Theory]
    [InlineData("../escape.dll")]
    [InlineData("runtime/x64/../../escape.dll")]
    [InlineData("runtime/x64/file.dll:stream")]
    public void ManifestCannotEscapeItsScope(string path)
    {
        var manifest = Fixture().Append("x64\t" + path + "\t" + new string('A', 64));
        Assert.Throws<IOException>(() => StartupPayload.Verify(root, "x64", manifest, _ => Assert.Fail("Invalid manifest")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RegistrationRestoresEnvironmentOnSuccessAndFailure(bool fail)
    {
        var names = new[] { "DOTNET_ROOT", "DOTNET_ROOT_X86", "DOTNET_ROOT_X64" };
        var before = names.Select(Environment.GetEnvironmentVariable).ToArray();
        var called = false;
        using var payload = StartupPayload.Verify(root, "x64", Fixture(), _ => called = true);
        Assert.True(called);
        bool Register(string path)
        {
            Assert.Equal(Path.Combine(root, "Nadabas.Loader.x64.xll"), path);
            Assert.All(names, name => Assert.Equal(Path.Combine(root, "runtime", "x64"), Environment.GetEnvironmentVariable(name)));
            if (fail) throw new IOException("Simulated Excel rejection");
            return true;
        }
        if (fail) Assert.Throws<IOException>(() => payload.Register(Register));
        else Assert.True(payload.Register(Register));
        Assert.Equal(before, names.Select(Environment.GetEnvironmentVariable).ToArray());
    }

    public void Dispose()
    {
        if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
    }
}
