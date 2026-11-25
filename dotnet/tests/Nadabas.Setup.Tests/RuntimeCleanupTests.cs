using System.Security.Cryptography;
using System.Text.Json;
using Nadabas.PackageUpdate;

namespace Nadabas.Setup.Tests;

public sealed class RuntimeCleanupTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "nadabas-cleanup-" + Guid.NewGuid().ToString("N"));
    private readonly string id = Guid.NewGuid().ToString("N");
    private string Root => Path.Combine(directory, "installation");
    private string Cache => Root + "-package-cache" + Path.DirectorySeparatorChar + id;
    private string Runtime => Path.Combine(Cache, "runtime", "x86");

    private string[] Fixture(string status = "installed")
    {
        Directory.CreateDirectory(Root);
        Directory.CreateDirectory(Cache);
        File.WriteAllText(Path.Combine(Cache, "package-update.json"), JsonSerializer.Serialize(new PendingPackage(1, Root, "1.0.0", "1.0.1", "unused", true)));
        File.WriteAllText(Path.Combine(Cache, "status.json"), JsonSerializer.Serialize(new { status }));
        File.WriteAllText(Path.Combine(Cache, "setup.exe"), "preserve installer");
        File.WriteAllText(Path.Combine(Cache, "installer.log"), "preserve diagnostics");
        return new[] { "dotnet.exe", "host/fxr/10.0.11/hostfxr.dll", "shared/Microsoft.NETCore.App/10.0.11/coreclr.dll",
            "shared/Microsoft.WindowsDesktop.App/10.0.11/System.Windows.Forms.dll" }.Select(relative =>
        {
            var path = Path.Combine(Runtime, relative);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, relative);
            return "x86\t" + relative + "\t" + Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
        }).ToArray();
    }

    [Fact]
    public void PreviewPreservesFilesAndApplyRetainsInstallerAndLogs()
    {
        if (!OperatingSystem.IsWindows()) return;
        var manifest = Fixture();
        Assert.Equal("reviewed", RuntimeCleanup.Run(Root, id, "x86", false, manifest).Status);
        Assert.True(File.Exists(Path.Combine(Runtime, "dotnet.exe")));
        Assert.Equal("removed", RuntimeCleanup.Run(Root, id, "x86", true, manifest).Status);
        Assert.False(Directory.Exists(Runtime));
        Assert.Equal("preserve installer", File.ReadAllText(Path.Combine(Cache, "setup.exe")));
        Assert.Equal("preserve diagnostics", File.ReadAllText(Path.Combine(Cache, "installer.log")));
        Assert.Equal("already-absent", RuntimeCleanup.Run(Root, id, "x86", true, manifest).Status);
    }

    [Theory]
    [InlineData("ready")]
    [InlineData("repair-required")]
    [InlineData("unknown")]
    public void UnfinishedOrFailedJobsArePreserved(string status)
    {
        if (!OperatingSystem.IsWindows()) return;
        var manifest = Fixture(status);
        Assert.Throws<IOException>(() =>
        {
            if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException();
            RuntimeCleanup.Run(Root, id, "x86", true, manifest);
        });
        Assert.True(File.Exists(Path.Combine(Runtime, "dotnet.exe")));
    }

    [Theory]
    [InlineData("pending")]
    [InlineData("active")]
    [InlineData("changed")]
    [InlineData("extra")]
    public void UnsafeRemovalIsRefusedBeforeDeletingFiles(string reason)
    {
        if (!OperatingSystem.IsWindows()) return;
        var manifest = Fixture();
        if (reason == "pending") File.WriteAllText(Path.Combine(Root, "pending-platform-update"), id);
        if (reason == "changed") File.AppendAllText(Path.Combine(Runtime, "dotnet.exe"), "changed");
        if (reason == "extra") File.WriteAllText(Path.Combine(Runtime, "unexpected.txt"), "preserve");
        using var active = reason == "active" ? RuntimeCache.Open(Cache, "x86", manifest) : null;
        Assert.Throws<IOException>(() =>
        {
            if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException();
            RuntimeCleanup.Run(Root, id, "x86", true, manifest);
        });
        Assert.True(File.Exists(Path.Combine(Runtime, "dotnet.exe")));
    }

    [Fact]
    public void WrongJobIdentityAndLockedFilesPreserveTheWholeRuntime()
    {
        if (!OperatingSystem.IsWindows()) return;
        var manifest = Fixture();
        var job = Path.Combine(Cache, "package-update.json");
        var original = File.ReadAllText(job);
        File.WriteAllText(job, JsonSerializer.Serialize(new PendingPackage(1, directory, "1.0.0", "1.0.1", "unused", true)));
        Assert.Throws<IOException>(() =>
        {
            if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException();
            RuntimeCleanup.Run(Root, id, "x86", true, manifest);
        });
        File.WriteAllText(job, original);
        using var locked = File.OpenRead(Path.Combine(Runtime, "host/fxr/10.0.11/hostfxr.dll"));
        Assert.Throws<IOException>(() =>
        {
            if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException();
            RuntimeCleanup.Run(Root, id, "x86", true, manifest);
        });
        Assert.Equal(4, Directory.GetFiles(Runtime, "*", SearchOption.AllDirectories).Length);
    }

    [Fact]
    public void InterruptedRemovalCanResume()
    {
        if (!OperatingSystem.IsWindows()) return;
        var manifest = Fixture("restored");
        File.Delete(Path.Combine(Runtime, "dotnet.exe"));
        Assert.Equal(3, RuntimeCleanup.Run(Root, id, "x86", true, manifest).Files);
    }

    [Fact]
    public void CancelledJobCanReleaseRuntimeButAttemptedInstallCannot()
    {
        if (!OperatingSystem.IsWindows()) return;
        var manifest = Fixture("ready");
        File.WriteAllText(Path.Combine(Cache, "cancelled.json"), "{}");
        File.WriteAllText(Path.Combine(Cache, "apply-started.json"), "{}");
        Assert.Throws<IOException>(() =>
        {
            if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException();
            RuntimeCleanup.Run(Root, id, "x86", true, manifest);
        });
        File.Delete(Path.Combine(Cache, "apply-started.json"));
        Assert.Equal("removed", RuntimeCleanup.Run(Root, id, "x86", true, manifest).Status);
    }

    public void Dispose()
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
    }
}
