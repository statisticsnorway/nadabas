using System.Diagnostics;
using System.Runtime.InteropServices;
using Nadabas.Hosting;
using Nadabas.PackageUpdate;

namespace Nadabas.Setup.Tests;

public sealed class RuntimeChildProcessTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), "nadabas-child-" + Guid.NewGuid());

    private ProcessStartInfo Helper(bool x64)
    {
        Directory.CreateDirectory(root);
        var path = Path.Combine(root, "helper.exe");
        using var stream = File.Create(path);
        using var writer = new BinaryWriter(stream);
        stream.SetLength(1024);
        writer.Write((ushort)0x5a4d);
        stream.Position = 0x3c; writer.Write(0x80);
        stream.Position = 0x80; writer.Write(0x4550);
        writer.Write((ushort)(x64 ? 0x8664 : 0x14c)); writer.Write((ushort)1);
        stream.Position = 0x94; writer.Write((ushort)(x64 ? 240 : 224)); writer.Write((ushort)2);
        writer.Write((ushort)(x64 ? 0x20b : 0x10b));
        return new(path) { UseShellExecute = false };
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PrivateRuntimeIsPassedToMatchingChildOnly(bool x64)
    {
        var arch = x64 ? "x64" : "x86";
        var start = Helper(x64);
        var before = Environment.GetEnvironmentVariable("DOTNET_ROOT");
        start.Environment["DOTNET_ROOT"] = "untrusted inherited value";
        var runtime = Path.Combine(root, "runtime", arch);
        var selected = RuntimeChildProcess.Configure(start, root,
            Path.Combine(runtime, "shared", "Microsoft.NETCore.App", "10.0.11"), x64 ? Architecture.X64 : Architecture.X86);
        Assert.Equal(runtime, selected);
        Assert.Equal(runtime, start.Environment["DOTNET_ROOT"]);
        Assert.Equal(runtime, start.Environment["DOTNET_ROOT_" + arch.ToUpperInvariant()]);
        Assert.Equal(before, Environment.GetEnvironmentVariable("DOTNET_ROOT"));
    }

    [Fact]
    public void DifferentHelperArchitectureFailsBeforeEnvironmentChange()
    {
        var start = Helper(false);
        start.Environment["DOTNET_ROOT"] = "original";
        Assert.Throws<IOException>(() => RuntimeChildProcess.Configure(start, root,
            Path.Combine(root, "runtime", "x64", "shared", "Microsoft.NETCore.App", "10.0.11"), Architecture.X64));
        Assert.Equal("original", start.Environment["DOTNET_ROOT"]);
    }

    [Fact]
    public void SharedRuntimeIsExplicitInsteadOfAmbientRoot()
    {
        var start = Helper(true);
        var runtime = Path.Combine(root + "-global", "dotnet");
        Assert.Equal(runtime, RuntimeChildProcess.Configure(start, root,
            Path.Combine(runtime, "shared", "Microsoft.NETCore.App", "10.0.11"), Architecture.X64));
        PackageRules.RequireOutside(root, runtime);
    }

    [Fact]
    public void WholePackageReplacementCannotRunOverItsRuntime()
    {
        var start = Helper(true);
        var runtime = RuntimeChildProcess.Configure(start, root,
            Path.Combine(root, "runtime", "x64", "shared", "Microsoft.NETCore.App", "10.0.11"), Architecture.X64);
        Assert.Throws<IOException>(() => PackageRules.RequireOutside(root, runtime));
    }

    [Fact]
    public async Task DirectApplyRejectsRuntimeReplacementBeforeMarkingJobStarted()
    {
        if (!OperatingSystem.IsWindows()) return;
        var cache = Path.Combine(root, "job");
        Directory.CreateDirectory(cache);
        var installation = RuntimeEnvironment.GetRuntimeDirectory();
        var job = new PendingPackage(1, installation, "1.0.0", "1.0.1", "unused", true);
        File.WriteAllText(Path.Combine(cache, "package-update.json"), System.Text.Json.JsonSerializer.Serialize(job));

        await Assert.ThrowsAsync<IOException>(() =>
        {
            if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException();
            return PackageJob.Apply(cache);
        });

        Assert.False(File.Exists(Path.Combine(cache, "apply.lock")));
        Assert.False(File.Exists(Path.Combine(cache, "apply-started.json")));
        Assert.False(File.Exists(Path.Combine(cache, "status.json")));
    }

    [Fact]
    public void PrivateRuntimeInWrongArchitectureDirectoryIsRejected()
    {
        var start = Helper(true);
        Assert.Throws<IOException>(() => RuntimeChildProcess.Configure(start, root,
            Path.Combine(root, "runtime", "x86", "shared", "Microsoft.NETCore.App", "10.0.11"), Architecture.X64));
    }

    [Fact]
    public void ShellExecutionCannotSilentlyDropEnvironment()
    {
        var start = Helper(true);
        start.UseShellExecute = true;
        Assert.Throws<IOException>(() => RuntimeChildProcess.Configure(start, root,
            Path.Combine(root, "runtime", "x64", "shared", "Microsoft.NETCore.App", "10.0.11"), Architecture.X64));
    }

    public void Dispose()
    {
        if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
    }
}
