using System.Security.Cryptography;
using Nadabas.PackageUpdate;

namespace Nadabas.Setup.Tests;

public sealed class RuntimeCacheTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "nadabas-runtime-cache-" + Guid.NewGuid().ToString("N"));
    private string Installation => Path.Combine(directory, "installation");
    private string Cache => Path.Combine(directory, "job");
    private const string Version = "10.0.11";

    private string[] Fixture(string architecture)
    {
        Directory.CreateDirectory(Cache);
        return new[] { "dotnet.exe", $"host/fxr/{Version}/hostfxr.dll",
            $"shared/Microsoft.NETCore.App/{Version}/coreclr.dll",
            $"shared/Microsoft.WindowsDesktop.App/{Version}/System.Windows.Forms.dll" }.Select(relative =>
        {
            var path = Path.Combine(Installation, "runtime", architecture, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, architecture + relative);
            return architecture + "\t" + relative + "\t" + Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
        }).ToArray();
    }

    [Theory]
    [InlineData("x86")]
    [InlineData("x64")]
    public void VerifiedCopyIsOutsideInstallationAndCanBeReused(string architecture)
    {
        var manifest = Fixture(architecture);
        using (var staged = RuntimeCache.Stage(Installation, Cache, architecture, manifest))
        {
            PackageRules.RequireOutside(Installation, staged.DirectoryPath);
            Assert.Equal(4, Directory.GetFiles(staged.DirectoryPath, "*", SearchOption.AllDirectories).Length);
        }
        using var retry = RuntimeCache.Stage(Installation, Cache, architecture, manifest);
        Assert.Equal(Path.Combine(Cache, "runtime", architecture), retry.DirectoryPath);
    }

    [Theory]
    [InlineData("changed")]
    [InlineData("extra")]
    [InlineData("missing")]
    public void InvalidSourceDoesNotProduceReadyRuntime(string failure)
    {
        var manifest = Fixture("x86");
        var file = Path.Combine(Installation, "runtime", "x86", "dotnet.exe");
        if (failure == "changed") File.AppendAllText(file, "changed");
        if (failure == "extra") File.WriteAllText(file + ".extra", "unexpected");
        if (failure == "missing") File.Delete(file);
        Assert.Throws<IOException>(() => RuntimeCache.Stage(Installation, Cache, "x86", manifest));
        Assert.False(Directory.Exists(Path.Combine(Cache, "runtime", "x86")));
    }

    [Fact]
    public void ChangedCacheIsRejectedWithoutOverwritingIt()
    {
        var manifest = Fixture("x64");
        using (RuntimeCache.Stage(Installation, Cache, "x64", manifest)) { }
        var file = Path.Combine(Cache, "runtime", "x64", "dotnet.exe");
        File.WriteAllText(file, "changed");
        Assert.Throws<IOException>(() => RuntimeCache.Stage(Installation, Cache, "x64", manifest));
        Assert.Equal("changed", File.ReadAllText(file));
    }

    [Fact]
    public void ActiveCopyBlocksAnotherWriter()
    {
        var manifest = Fixture("x86");
        using var gate = new FileStream(Path.Combine(Cache, "runtime-stage.lock"), FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        Assert.Throws<IOException>(() => RuntimeCache.Stage(Installation, Cache, "x86", manifest));
    }

    [Fact]
    public void RetainedReadLeasesPreventReplacementOnWindows()
    {
        if (!OperatingSystem.IsWindows()) return;
        var manifest = Fixture("x64");
        string file;
        using (var staged = RuntimeCache.Stage(Installation, Cache, "x64", manifest))
        {
            file = Path.Combine(staged.DirectoryPath, "dotnet.exe");
            Assert.Throws<IOException>(() => File.WriteAllText(file, "changed"));
        }
        File.WriteAllText(file, "released");
    }

    [Fact]
    public void ChildLeaseSurvivesParentLeaseRelease()
    {
        var manifest = Fixture("x86");
        using var parent = RuntimeCache.Stage(Installation, Cache, "x86", manifest);
        using var child = RuntimeCache.LeaseForExecution(Installation, Cache, parent.CoreDirectory,
            System.Runtime.InteropServices.Architecture.X86, manifest);
        Assert.NotNull(child);
        parent.Dispose();
        var file = Path.Combine(child.DirectoryPath, "dotnet.exe");
        if (OperatingSystem.IsWindows()) Assert.Throws<IOException>(() => File.WriteAllText(file, "locked"));
        child.Dispose();
        File.WriteAllText(file, "released");
    }

    [Fact]
    public void WrongLoadedRuntimeIsRejectedAndReleasesItsLeases()
    {
        var manifest = Fixture("x86");
        using (RuntimeCache.Stage(Installation, Cache, "x86", manifest)) { }
        Assert.Throws<IOException>(() => RuntimeCache.LeaseForExecution(Installation, Cache,
            Path.Combine(directory, "shared-runtime", "shared", "Microsoft.NETCore.App", Version),
            System.Runtime.InteropServices.Architecture.X86, manifest));
        File.WriteAllText(Path.Combine(Cache, "runtime", "x86", "dotnet.exe"), "released after rejection");
    }

    [Fact]
    public void DeletedCacheCannotFallBackToSharedRuntimeRules()
    {
        var manifest = Fixture("x64");
        var expectedLoaded = Path.Combine(Cache, "runtime", "x64", "shared", "Microsoft.NETCore.App", Version);
        Assert.Throws<DirectoryNotFoundException>(() => RuntimeCache.LeaseForExecution(Installation, Cache,
            expectedLoaded, System.Runtime.InteropServices.Architecture.X64, manifest));
    }

    [Fact]
    public void LegacySharedRuntimeNeedsNoPrivateManifest()
    {
        Directory.CreateDirectory(Cache);
        using var lease = RuntimeCache.LeaseForExecution(Installation, Cache,
            Path.Combine(directory, "shared-runtime", "shared", "Microsoft.NETCore.App", Version),
            System.Runtime.InteropServices.Architecture.X64, []);
        Assert.Null(lease);
    }

    [Fact]
    public void OpenDoesNotRepairTamperedRuntimeBeforeLaunch()
    {
        var manifest = Fixture("x86");
        using (RuntimeCache.Stage(Installation, Cache, "x86", manifest)) { }
        File.WriteAllText(Path.Combine(Cache, "runtime", "x86", "dotnet.exe"), "tampered");
        Assert.Throws<IOException>(() => RuntimeCache.Open(Cache, "x86", manifest));
    }

    [Fact]
    public void CancellationLeavesNoReadyCopyAndAllowsRetry()
    {
        var manifest = Fixture("x86");
        Assert.Throws<OperationCanceledException>(() => RuntimeCache.Stage(Installation, Cache, "x86", manifest, new CancellationToken(true)));
        Assert.False(Directory.Exists(Path.Combine(Cache, "runtime", "x86")));
        using var retry = RuntimeCache.Stage(Installation, Cache, "x86", manifest);
    }

    [Fact]
    public void InterruptedCopyIsNotReusedAsReady()
    {
        var manifest = Fixture("x86");
        var partial = Path.Combine(Cache, "runtime-pending-interrupted");
        Directory.CreateDirectory(partial);
        File.WriteAllText(Path.Combine(partial, "dotnet.exe"), "partial");
        using var staged = RuntimeCache.Stage(Installation, Cache, "x86", manifest);
        Assert.Equal("partial", File.ReadAllText(Path.Combine(partial, "dotnet.exe")));
        Assert.NotEqual(partial, staged.DirectoryPath);
    }

    [Theory]
    [InlineData("../escape")]
    [InlineData("host/fxr/10.0.11/../escape")]
    [InlineData("host/fxr/10.0.11/file:stream")]
    public void UnsafeManifestPathsAreRejected(string path)
    {
        var manifest = Fixture("x64").Append("x64\t" + path + "\t" + new string('A', 64));
        Assert.Throws<IOException>(() => RuntimeCache.Stage(Installation, Cache, "x64", manifest));
    }

    [Fact]
    public void MissingEmbeddedApprovalAndWrongArchitectureAreRejected()
    {
        var manifest = Fixture("x86");
        Assert.Throws<IOException>(() => RuntimeCache.Stage(Installation, Cache, "x64", manifest));
        Assert.Throws<IOException>(() => RuntimeCache.Stage(Installation, Cache, "x86"));
        Assert.Throws<IOException>(() => RuntimeCache.Stage(Installation, Installation, "x86", manifest));
    }

    public void Dispose()
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
    }
}
