using System.Security.Cryptography;
using System.Text.Json;
using Nadabas.Versioning;

namespace Nadabas.Setup.Tests;

public sealed class PendingUpdateTests : IDisposable
{
    private readonly string parent = Path.Combine(Path.GetTempPath(), "nadabas-pending-" + Guid.NewGuid().ToString("N"));
    private readonly string root;
    private readonly VersionStore store;
    private string PendingPath => Path.Combine(root, "pending-update.json");

    public PendingUpdateTests()
    {
        Directory.CreateDirectory(parent);
        root = Path.Combine(parent, "store");
        var oldFile = Path.Combine(parent, "old.xlam");
        var newFile = Path.Combine(parent, "new.xlam");
        File.WriteAllText(oldFile, "Synthetic old release");
        File.WriteAllText(newFile, "Synthetic new release");
        store = VersionStore.Create(root, new Dictionary<string, ApprovedRelease>
        {
            ["slot-old"] = new("6.02.005", Hash(oldFile)),
            ["slot-new"] = new("6.02.006", Hash(newFile))
        });
        store.Stage("slot-old", oldFile);
        store.Stage("slot-new", newFile);
        store.Activate("slot-old");
    }

    [Fact]
    public void ReadyNotificationRevalidatesWithoutActivatingAndDisappearsAfterCancellation()
    {
        Assert.Null(store.GetReadyUpdate());
        store.QueueActivation("slot-new");
        var before = File.ReadAllText(Path.Combine(root, "active.json"));
        Assert.Equal(new ReadyUpdate("slot-new", "6.02.006"), store.GetReadyUpdate());
        Assert.Equal(before, File.ReadAllText(Path.Combine(root, "active.json")));
        Assert.True(File.Exists(PendingPath));
        store.CancelPendingActivation();
        Assert.Null(store.GetReadyUpdate());
    }

    [Fact]
    public void NotificationDoesNotClaimTamperedUpdateIsReady()
    {
        store.QueueActivation("slot-new");
        var payload = Path.Combine(root, "versions", "slot-new", "NADABAS.TEST.slot-new.xlam");
        File.SetAttributes(payload, FileAttributes.Normal);
        File.WriteAllText(payload, "tampered");
        Assert.Throws<IOException>(() => store.GetReadyUpdate());
        using var selected = store.Select();
        Assert.Equal("slot-old", selected.Slot);
    }

    [Fact]
    public void QueueDoesNotSwitchUntilStartupAndExistingSessionKeepsOldRelease()
    {
        using var oldSession = store.Select();
        store.QueueActivation("slot-new");
        using (var inspected = store.Select()) Assert.Equal("slot-old", inspected.Slot);
        var startup = store.SelectForStartup();
        using var newSession = startup.Selection;
        Assert.Equal("applied", startup.UpdateStatus);
        Assert.Equal("6.02.006", newSession.ReleaseVersion);
        Assert.Equal("6.02.005", oldSession.ReleaseVersion);
        Assert.False(File.Exists(PendingPath));
        var nextStartup = store.SelectForStartup();
        using var next = nextStartup.Selection;
        Assert.Equal("none", nextStartup.UpdateStatus);
        Assert.Equal("slot-new", next.Slot);
    }

    [Fact]
    public void ModifiedPayloadAfterQueueIsRejectedWithoutChangingActivePointer()
    {
        store.QueueActivation("slot-new");
        var pointerHash = Hash(Path.Combine(root, "active.json"));
        var payload = Path.Combine(root, "versions", "slot-new", "NADABAS.TEST.slot-new.xlam");
        File.SetAttributes(payload, FileAttributes.Normal);
        File.WriteAllText(payload, "tampered after validation");
        var startup = store.SelectForStartup();
        using var selected = startup.Selection;
        Assert.Equal("rejected", startup.UpdateStatus);
        Assert.NotNull(startup.UpdateError);
        Assert.Equal("slot-old", selected.Slot);
        Assert.Equal(pointerHash, Hash(Path.Combine(root, "active.json")));
    }

    [Theory]
    [InlineData("not-json")]
    [InlineData("{\"SchemaVersion\":2,\"PreviousSlot\":\"slot-old\",\"TargetSlot\":\"slot-new\"}")]
    [InlineData("{\"SchemaVersion\":1,\"PreviousSlot\":\"slot-old\",\"TargetSlot\":\"../outside\"}")]
    [InlineData("{\"SchemaVersion\":1,\"PreviousSlot\":\"slot-stale\",\"TargetSlot\":\"slot-new\"}")]
    public void InvalidPendingRecordLeavesUsableOldVersion(string json)
    {
        File.WriteAllText(PendingPath, json);
        var startup = store.SelectForStartup();
        using var selected = startup.Selection;
        Assert.Equal("rejected", startup.UpdateStatus);
        Assert.Equal("slot-old", selected.Slot);
    }

    [Fact]
    public void CrashAfterPointerSwitchBeforePendingCleanupIsIdempotent()
    {
        var pending = store.QueueActivation("slot-new");
        var first = store.SelectForStartup();
        first.Selection.Dispose();
        File.WriteAllText(PendingPath, JsonSerializer.Serialize(pending));
        var recovered = store.SelectForStartup();
        using var selected = recovered.Selection;
        Assert.Equal("already-applied", recovered.UpdateStatus);
        Assert.Equal("slot-new", selected.Slot);
        Assert.False(File.Exists(PendingPath));
    }

    [Fact]
    public void CancellationAndExplicitActivationPreventSurpriseLaterUpdate()
    {
        store.QueueActivation("slot-new");
        store.CancelPendingActivation();
        Assert.False(File.Exists(PendingPath));
        using (var selected = store.Select()) Assert.Equal("slot-old", selected.Slot);
        store.QueueActivation("slot-new");
        store.Activate("slot-old");
        Assert.False(File.Exists(PendingPath));
        var startup = store.SelectForStartup();
        using var same = startup.Selection;
        Assert.Equal("none", startup.UpdateStatus);
        Assert.Equal("slot-old", same.Slot);
    }

    [Fact]
    public void DowngradeOrSameVersionCannotBeQueuedAsAutomaticUpgrade()
    {
        Assert.Throws<IOException>(() => store.QueueActivation("slot-old"));
        store.Activate("slot-new");
        Assert.Throws<IOException>(() => store.QueueActivation("slot-old"));
        Assert.False(File.Exists(PendingPath));
    }

    private static string Hash(string file) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(file)));
    public void Dispose()
    {
        foreach (var file in Directory.EnumerateFiles(parent, "*", SearchOption.AllDirectories)) File.SetAttributes(file, FileAttributes.Normal);
        Directory.Delete(parent, true);
    }
}
