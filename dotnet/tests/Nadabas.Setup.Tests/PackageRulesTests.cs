using Nadabas.PackageUpdate;

namespace Nadabas.Setup.Tests;

public sealed class PackageRulesTests
{
    [Theory]
    [InlineData("0.7.1", "0.7.1.0")]
    [InlineData("0.8.0", "0.7.1")]
    public void SameOrOlderPackageIsNotAnAutomaticUpgrade(string current, string incoming)
        => Assert.Throws<IOException>(() => PackageRules.RequireUpgrade(current, incoming));

    [Fact]
    public void PackageVersionCanAdvanceIndependentlyOfVbaVersion()
        => PackageRules.RequireUpgrade("0.7.1", "0.8.0.0");

    [Fact]
    public void RunningHelperAndCacheCannotBeInsideReplacedInstallation()
    {
        var root = Path.Combine(Path.GetTempPath(), "nadabas-install");
        Assert.Throws<IOException>(() => PackageRules.RequireOutside(root, root));
        Assert.Throws<IOException>(() => PackageRules.RequireOutside(root, Path.Combine(root, "updater", "helper.exe")));
        PackageRules.RequireOutside(root, root + "-cache");
    }

    [Fact]
    public void RelativeJobPathsAreRejected()
        => Assert.Throws<IOException>(() => PackageRules.SafePath("relative/cache"));

    [Fact]
    public void CachedHelperAllowsWindowsPathAliasButRejectsIdenticalCopy()
    {
        if (!OperatingSystem.IsWindows()) return;
        var file = Path.GetTempFileName();
        var copy = file + ".copy";
        try
        {
            File.Copy(file, copy);
            PackageRules.RequireCachedExecutable(Path.Join(@"\\?\", file), file);
            Assert.Throws<IOException>(() => PackageRules.RequireCachedExecutable(copy, file));
        }
        finally
        {
            File.Delete(copy);
            File.Delete(file);
        }
    }
}
