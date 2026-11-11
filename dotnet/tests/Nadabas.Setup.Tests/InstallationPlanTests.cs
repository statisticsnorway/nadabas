using System.Text.Json;
using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class InstallationPlanTests : IDisposable
{
    private readonly string root = Path.Combine(Path.GetTempPath(), "nadabas-plan-" + Guid.NewGuid());
    private readonly string source;
    private readonly string target;

    public InstallationPlanTests()
    {
        Directory.CreateDirectory(root);
        source = Path.Combine(root, "NADABAS.6.02.004.xlam");
        target = Path.Combine(root, "target with spaces æøå");
        // Deliberately not an Office file: a plan must never claim package/signature validation.
        File.WriteAllText(source, "candidate bytes");
    }

    [Fact]
    public void FreshInstallIsOnlyAProposalAndCreatesNothing()
    {
        var plan = InstallationPlanner.Create(source, target);
        Assert.Equal("review-required", plan.Status);
        Assert.False(plan.CanApply);
        Assert.Equal("install-file", Assert.Single(plan.Steps).Action);
        Assert.Equal(64, plan.Source!.Sha256.Length);
        Assert.False(Directory.Exists(target));
        Assert.Equal("candidate bytes", File.ReadAllText(source));
    }

    [Fact]
    public void ReplacementRequiresBackupAndPreservesAllFiles()
    {
        Directory.CreateDirectory(target);
        var installed = Path.Combine(target, "NADABAS.XLAM");
        File.WriteAllText(installed, "old bytes");
        File.WriteAllText(Path.Combine(target, "unrelated.xlam"), "other");
        var plan = InstallationPlanner.Create(source, target);
        Assert.Equal(new[] { "backup-required-before-replacement", "replace-file" }, plan.Steps.Select(s => s.Action));
        Assert.Contains(plan.ExistingTarget!.Sha256, plan.Steps[0].Destination);
        Assert.Equal("old bytes", File.ReadAllText(installed));
        Assert.Equal("other", File.ReadAllText(Path.Combine(target, "unrelated.xlam")));
        Assert.False(Directory.Exists(Path.Combine(target, ".nadabas-backups")));
    }

    [Fact]
    public void IdenticalContentDoesNotProposeReplacement()
    {
        Directory.CreateDirectory(target);
        File.Copy(source, Path.Combine(target, "NADABAS.XLAM"));
        Assert.Equal("keep-identical-file", Assert.Single(InstallationPlanner.Create(source, target).Steps).Action);
    }

    [Fact]
    public void SourceAlreadyAtDestinationIsKept()
    {
        Directory.CreateDirectory(target);
        var installed = Path.Combine(target, "NADABAS.XLAM");
        File.Copy(source, installed);
        Assert.Equal("keep-identical-file", Assert.Single(InstallationPlanner.Create(installed, target).Steps).Action);
    }

    [Fact]
    public void OldVersionedAddinBlocksPlanWithoutDeletion()
    {
        Directory.CreateDirectory(target);
        var old = Path.Combine(target, "NADABAS.6.01.xlam");
        File.WriteAllText(old, "old");
        var plan = InstallationPlanner.Create(source, target);
        Assert.Equal("blocked", plan.Status);
        Assert.Empty(plan.Steps);
        Assert.Equal("old", File.ReadAllText(old));
    }

    [Fact]
    public void MissingSourceBlocksPlan()
    {
        var plan = InstallationPlanner.Create(Path.Combine(root, "NADABAS.missing.xlam"), target);
        Assert.Equal("blocked", plan.Status);
        Assert.Empty(plan.Steps);
        Assert.False(Directory.Exists(target));
    }

    [Fact]
    public void EmptySourceBlocksPlan()
    {
        File.WriteAllText(source, "");
        Assert.Equal("blocked", InstallationPlanner.Create(source, target).Status);
    }

    [Fact]
    public void FileAsTargetDirectoryBlocksPlan()
    {
        File.WriteAllText(target, "keep");
        Assert.Equal("blocked", InstallationPlanner.Create(source, target).Status);
        Assert.Equal("keep", File.ReadAllText(target));
    }

    [Fact]
    public void DirectoryAtCanonicalDestinationBlocksPlan()
    {
        Directory.CreateDirectory(Path.Combine(target, "NADABAS.XLAM"));
        Assert.Equal("blocked", InstallationPlanner.Create(source, target).Status);
    }

    [Fact]
    public void FileInTargetAncestorsBlocksPlan()
    {
        File.WriteAllText(target, "keep");
        Assert.Equal("blocked", InstallationPlanner.Create(source, Path.Combine(target, "child")).Status);
    }

    [Fact]
    public void WrongSourceNameBlocksPlan()
    {
        var wrong = Path.Combine(root, "other.xlam");
        File.WriteAllText(wrong, "bytes");
        var plan = InstallationPlanner.Create(wrong, target);
        Assert.Equal("blocked", plan.Status);
        Assert.Empty(plan.Steps);
    }

    [Fact]
    public void BlockedPlanReturnsFailureWithJson()
    {
        var output = new StringWriter();
        Assert.Equal(1, Command.Run(["plan-install", "--source", source + ".missing", "--addin-directory", target], output, new StringWriter()));
        using var json = JsonDocument.Parse(output.ToString());
        Assert.Equal("blocked", json.RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void CommandReturnsDistinctPlanContract()
    {
        var output = new StringWriter();
        var error = new StringWriter();
        Assert.Equal(0, Command.Run(["plan-install", "--source", source, "--addin-directory", target], output, error));
        using var json = JsonDocument.Parse(output.ToString());
        Assert.Equal("installation-plan", json.RootElement.GetProperty("kind").GetString());
        Assert.False(json.RootElement.GetProperty("canApply").GetBoolean());
        Assert.Empty(error.ToString());
    }

    [Theory]
    [InlineData("relative.xlam", "relative")]
    [InlineData("", "")]
    public void InvalidPathsReturnUsageError(string input, string destination)
    {
        Assert.Equal(2, Command.Run(["plan-install", "--source", input, "--addin-directory", destination], new StringWriter(), new StringWriter()));
    }

    public void Dispose() => Directory.Delete(root, recursive: true);
}
