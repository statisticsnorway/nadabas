using System.Security.Cryptography;
using Nadabas.Versioning;

namespace Nadabas.Setup.Tests;

public sealed class VersionStoreTests : IDisposable
{
    private readonly string parent = Path.Combine(Path.GetTempPath(), "nadabas-store-tests-" + Guid.NewGuid().ToString("N"));
    private readonly string root;
    private readonly string source;
    private readonly VersionStore store;
    public VersionStoreTests()
    {
        Directory.CreateDirectory(parent);
        root = Path.Combine(parent, "store");
        source = Path.Combine(parent, "fixture.xlam");
        File.WriteAllText(source, "Synthetic version store fixture, not Excel content");
        store = VersionStore.Create(root, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(source))));
        store.Stage("slot-a", source);
        store.Stage("slot-b", source);
        store.Activate("slot-a");
    }

    [Fact]
    public void DistinctReleasesKeepTheirOwnHashVersionAndRollbackSelection()
    {
        var next = Path.Combine(parent, "next.xlam");
        File.WriteAllText(next, "Different simulated release content");
        var catalog = new Dictionary<string, ApprovedRelease>
        {
            ["slot-old"] = new("6.02.005", Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(source)))),
            ["slot-new"] = new("6.02.006", Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(next))))
        };
        var versions = VersionStore.Create(Path.Combine(parent, "distinct"), catalog);
        versions.Stage("slot-old", source);
        Assert.Throws<IOException>(() => versions.Stage("slot-new", source));
        versions.Stage("slot-new", next);
        versions.Activate("slot-old");
        using var oldSession = versions.Select();
        versions.Activate("slot-new");
        using var newSession = versions.Select();
        Assert.Equal("6.02.005", oldSession.ReleaseVersion);
        Assert.Equal("6.02.006", newSession.ReleaseVersion);
        versions.Activate("slot-old");
        using var rollback = versions.Select();
        Assert.Equal(oldSession.AddinPath, rollback.AddinPath);
        Assert.Equal("6.02.005", rollback.ReleaseVersion);
        Assert.Equal("6.02.006", newSession.ReleaseVersion);
        Assert.Throws<IOException>(() => versions.Stage("slot-unapproved", next));
    }

    [Fact]
    public void ExistingSessionKeepsItsSelectionAcrossActivationAndRollback()
    {
        using var first = store.Select();
        store.Activate("slot-b");
        using var second = store.Select();
        Assert.Equal("slot-a", first.Slot);
        Assert.Equal("slot-b", second.Slot);
        Assert.NotEqual(first.AddinPath, second.AddinPath);
        store.Activate("slot-a");
        using var third = store.Select();
        Assert.Equal(first.AddinPath, third.AddinPath);
        Assert.Equal("slot-b", second.Slot);
        Assert.True(File.Exists(first.AddinPath));
        Assert.True(File.Exists(second.AddinPath));
    }

    [Theory]
    [InlineData("../outside")]
    [InlineData("slot-a/../../outside")]
    [InlineData("C:\\outside")]
    [InlineData("slot-missing")]
    public void InvalidActivationPreservesPreviousPointer(string slot)
    {
        var before = File.ReadAllText(Path.Combine(root, "active.json"));
        Assert.ThrowsAny<IOException>(() => store.Activate(slot));
        Assert.Equal(before, File.ReadAllText(Path.Combine(root, "active.json")));
    }

    [Fact]
    public void ExistingSlotCannotBeOverwritten()
    {
        Assert.Throws<IOException>(() => store.Stage("slot-a", source));
        using var selected = store.Select();
        Assert.Equal(File.ReadAllBytes(source), File.ReadAllBytes(selected.AddinPath));
    }

    [Fact]
    public void TamperedVersionIsRejectedBeforePointerChanges()
    {
        var path = Path.Combine(root, "versions", "slot-b", "NADABAS.TEST.slot-b.xlam");
        File.SetAttributes(path, FileAttributes.Normal);
        File.WriteAllText(path, "changed");
        Assert.Throws<IOException>(() => store.Activate("slot-b"));
        using var selected = store.Select();
        Assert.Equal("slot-a", selected.Slot);
    }

    [Fact]
    public void WindowsSessionLeasePreventsDeletionEvenIfAttributesAreCleared()
    {
        if (!OperatingSystem.IsWindows()) return;
        using var selected = store.Select();
        File.SetAttributes(selected.AddinPath, FileAttributes.Normal);
        Assert.Throws<IOException>(() => File.Delete(selected.AddinPath));
        Assert.Throws<IOException>(() => File.WriteAllText(selected.AddinPath, "replacement"));
    }

    [Fact]
    public void UnexpectedSchemaIsRejected()
    {
        File.WriteAllText(Path.Combine(root, "active.json"), "{\"SchemaVersion\":2,\"Slot\":\"slot-a\"}");
        Assert.Throws<IOException>(() => store.Select());
    }

    public void Dispose()
    {
        foreach (var file in Directory.EnumerateFiles(parent, "*", SearchOption.AllDirectories)) File.SetAttributes(file, FileAttributes.Normal);
        Directory.Delete(parent, recursive: true);
    }
}
