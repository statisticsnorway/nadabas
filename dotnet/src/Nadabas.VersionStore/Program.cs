using System.Text.Json;
using Nadabas.Versioning;

try
{
    if (args.Length == 4 && args[0] == "restart-excel" && Path.IsPathFullyQualified(args[1]) &&
        int.TryParse(args[2], out var processId) && long.TryParse(args[3], out var startTicks))
    {
        await ExcelRestart.Run(args[1], processId, startTicks);
        return 0;
    }
    if (args.Length == 3 && args[0] == "prepare" && Path.IsPathFullyQualified(args[1]) && Path.IsPathFullyQualified(args[2]))
    {
        var store = VersionStore.Create(args[1]);
        store.Stage("slot-a", args[2]);
        store.Stage("slot-b", args[2]);
        store.Activate("slot-a");
        Console.WriteLine("Prepared two test slots with identical approved 6.02.005 bytes.");
        return 0;
    }
    if (args.Length == 4 && args[0] == "stage" && Path.IsPathFullyQualified(args[1]) && Path.IsPathFullyQualified(args[3]))
    {
        new VersionStore(args[1]).Stage(args[2], args[3]);
        Console.WriteLine("Approved release staged: " + args[2]);
        return 0;
    }
    if (args.Length == 3 && args[0] == "queue" && Path.IsPathFullyQualified(args[1]))
    {
        var store = new VersionStore(args[1]);
        var pending = store.QueueActivation(args[2]);
        Console.WriteLine(JsonSerializer.Serialize(new { status = "ready-on-next-start", pending.PreviousSlot, pending.TargetSlot }));
        Console.WriteLine("A verified NADABAS update is ready. Save your work and close all Excel windows, then reopen Excel to finish updating.");
        return 0;
    }
    if (args.Length == 2 && args[0] == "cancel-update" && Path.IsPathFullyQualified(args[1]))
    {
        new VersionStore(args[1]).CancelPendingActivation();
        Console.WriteLine("Pending update cancelled. Active version unchanged.");
        return 0;
    }
    if (args.Length == 3 && args[0] == "activate" && Path.IsPathFullyQualified(args[1]))
    {
        new VersionStore(args[1]).Activate(args[2]);
        Console.WriteLine("Active for next session: " + args[2]);
        return 0;
    }
    if (args.Length == 2 && args[0] == "inspect" && Path.IsPathFullyQualified(args[1]))
    {
        using var selected = new VersionStore(args[1]).Select();
        Console.WriteLine(JsonSerializer.Serialize(new { selected.Slot, selected.AddinPath, selected.ReleaseVersion }));
        return 0;
    }
    Console.Error.WriteLine("TEST ONLY: prepare <new-absolute-root> <approved-xlam> | stage <absolute-root> <approved-slot> <approved-xlam> | activate <absolute-root> <approved-slot> | queue <absolute-root> <approved-slot> | cancel-update <absolute-root> | inspect <absolute-root>");
    return 2;
}
catch (Exception ex) { Console.Error.WriteLine(ex.Message); return 1; }
