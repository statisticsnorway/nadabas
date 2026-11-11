using ExcelDna.Integration;
using Nadabas.Versioning;
using Nadabas.PackageUpdate;
using System.Windows.Forms;

namespace Nadabas.ExcelBootstrap;

public static class UpdateCommands
{
    internal static UpdateNotifications? Current { get; set; }

    [ExcelCommand(Name = "NADABAS.CheckUpdates", Description = "Show NADABAS update status")]
    public static void CheckUpdates() => Current?.Poll(manual: true);
}

internal sealed class UpdateNotifications(string root) : IDisposable
{
    private readonly System.Windows.Forms.Timer timer = new() { Interval = 3000 };
    private Form? notice;
    private string? shownSlot;
    private bool queued;
    private bool disposed;
    private bool readyNotice;
    private string? shownLanguage;
    private string? checkingPlatform;
    private bool manualRequested;
    private string? lastPlatform;
    private DateTime nextVerification;
    private DateTime nextDownloadCheck;
    private bool checkingDownload;
    private bool downloadManualRequested;
    private readonly CancellationTokenSource lifetime = new();
    private string downloadStatus = "not-configured";
    private UpdateDiagnostic? downloadDiagnostic;

    internal void Start()
    {
        UpdateCommands.Current = this;
        timer.Tick += (_, _) =>
        {
            if (disposed || queued) return;
            queued = true;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                queued = false;
                if (!disposed) Poll(manual: false);
            });
        };
        timer.Start();
    }

    internal void Poll(bool manual, bool skipDownload = false)
    {
        if (disposed) return;
        var language = UpdateLanguage.CurrentCode();
        string T(string key) => UpdateLanguage.Text(language, key);
        try
        {
            var id = PlatformHandoff.ReadId(root);
            if (id is not null)
            {
                if (checkingPlatform is not null) { manualRequested |= manual; return; }
                if (!manual && lastPlatform == id && DateTime.UtcNow < nextVerification &&
                    (notice is null || shownLanguage == language)) return;
                checkingPlatform = id;
                manualRequested = manual;
                _ = CheckPlatformAsync(id);
                return;
            }
            lastPlatform = null;
            var ready = new VersionStore(root).GetReadyUpdate();
            if (ready is null)
            {
                if (!skipDownload && (manual || DateTime.UtcNow >= nextDownloadCheck))
                {
                    downloadManualRequested |= manual;
                    if (!checkingDownload)
                    {
                        checkingDownload = true;
                        _ = CheckDownloadAsync(manual);
                    }
                    if (manual) Show(language, T("status"), T("checking"), false);
                    return;
                }
                shownSlot = null;
                if (readyNotice) notice?.Close();
                if (manual || (notice is not null && shownLanguage != language)) Show(language, T("status"), T(DownloadMessageKey(downloadStatus)), false, diagnostic: downloadDiagnostic);
                return;
            }
            if (!manual && shownSlot == ready.Slot && (notice is null || shownLanguage == language)) return;
            shownSlot = ready.Slot;
            Show(language, string.Format(T("ready"), ready.Version), T("body") + "\n\n" + T("laterHint"), true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException or ArgumentException)
        {
            if (readyNotice) notice?.Close();
            shownSlot = null;
            var diagnostic = Record(UpdateStage.Verification, ex);
            if (manual) Show(language, T("verifyFailed"), T("platformReview"), false, diagnostic: diagnostic);
        }
    }

    private static string DownloadMessageKey(string status) => status switch
    {
        "not-configured" => "offline",
        "disabled" => "updatesDisabled",
        "failed" => "downloadFailed",
        "checking" => "checking",
        "skipped" => "updateSkipped",
        _ => "none"
    };

    private async Task CheckDownloadAsync(bool manual)
    {
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(lifetime.Token);
        deadline.CancelAfter(TimeSpan.FromMinutes(4));
        string status;
        UpdateDiagnostic? diagnostic = null;
        try
        {
            var result = await Task.Run(() => AutomaticUpdates.CheckFromExcel(root, manual, deadline.Token)).ConfigureAwait(false);
            status = result.Status;
            diagnostic = result.Diagnostic;
            diagnostic?.TrySave(root);
        }
        catch (Exception ex)
        {
            diagnostic = Record(UpdateStage.Download, ex);
            status = "failed";
        }
        if (disposed) return;
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                checkingDownload = false;
                if (disposed) return;
                manual |= downloadManualRequested;
                downloadManualRequested = false;
                downloadStatus = status;
                downloadDiagnostic = diagnostic;
                nextDownloadCheck = DateTime.UtcNow.Add(status == "checking" ? TimeSpan.FromSeconds(30) : TimeSpan.FromHours(6));
                Poll(manual, skipDownload: true);
            });
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
    }

    private async Task CheckPlatformAsync(string id)
    {
        PlatformOffer? offer = null;
        Exception? failure = null;
        try { offer = await Task.Run(() => PlatformHandoff.Verify(root, id)).ConfigureAwait(false); }
        catch (Exception ex) { failure = ex; }
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            checkingPlatform = null;
            if (disposed) return;
            var manual = manualRequested;
            manualRequested = false;
            var language = UpdateLanguage.CurrentCode();
            string T(string key) => UpdateLanguage.Text(language, key);
            try
            {
                if (PlatformHandoff.ReadId(root) != id) { Poll(manual); return; }
                lastPlatform = id;
                nextVerification = DateTime.UtcNow.AddSeconds(30);
                if (failure is not null || offer is null) throw new IOException("Package verification failed.", failure);
                var identity = "package:" + id;
                if (!manual && shownSlot == identity && (notice is null || shownLanguage == language)) return;
                shownSlot = identity;
                Show(language, string.Format(T("platformReady"), offer.Version), T("body") + "\n\n" + T("laterHint"), true, id);
            }
            catch (Exception ex)
            {
                var diagnostic = Record(UpdateStage.Verification, ex);
                shownSlot = null;
                if (readyNotice) notice?.Close();
                if (manual) Show(language, T("verifyFailed"), T("platformReview"), false, diagnostic: diagnostic);
            }
        });
    }

    private void Show(string language, string heading, string message, bool canCancel, string? platformId = null, UpdateDiagnostic? diagnostic = null)
    {
        string T(string key) => UpdateLanguage.Text(language, key);
        shownLanguage = language;
        notice?.Close();
        readyNotice = canCancel;
        var form = new Form
        {
            Text = T("title"),
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(16),
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            StartPosition = FormStartPosition.CenterScreen,
            AutoScaleMode = AutoScaleMode.Dpi
        };
        var layout = new TableLayoutPanel { AutoSize = true, ColumnCount = 1, Dock = DockStyle.Fill };
        form.Controls.Add(layout);
        layout.Controls.Add(new Label { Text = heading, AutoSize = true, Font = new System.Drawing.Font(System.Drawing.SystemFonts.MessageBoxFont!, System.Drawing.FontStyle.Bold) });
        layout.Controls.Add(new Label { Text = message, AutoSize = true, MaximumSize = new System.Drawing.Size(480, 0), Margin = new Padding(3, 14, 3, 16) });
        var actions = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        layout.Controls.Add(actions);
        var later = new Button { Text = canCancel ? T("later") : T("close"), AutoSize = true };
        later.Click += (_, _) => form.Close();
        actions.Controls.Add(later);
        form.CancelButton = later;
        if (diagnostic is not null)
        {
            var details = new Button { Text = T("details"), AutoSize = true };
            details.Click += (_, _) => ShowDetails(form, language, diagnostic);
            actions.Controls.Add(details);
        }
        if (canCancel)
        {
            var restart = new Button { Text = T("restart"), AutoSize = true };
            restart.Click += (_, _) =>
            {
                restart.Enabled = false;
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try { RestartUpdate.Start(root, language, platformId); }
                    catch (RestartBlockedException ex) { MessageBox.Show(ex.Message, T("title")); }
                    catch (Exception ex)
                    {
                        var diagnostic = Record(UpdateStage.Restart, ex);
                        MessageBox.Show(form, T("restartFailed"), T("title"));
                        ShowDetails(form, language, diagnostic);
                    }
                    finally { if (!restart.IsDisposed) restart.Enabled = true; }
                });
            };
            actions.Controls.Add(restart);
            form.AcceptButton = restart;
            var cancel = new Button { Text = T("cancel"), AutoSize = true };
            cancel.Click += (_, _) =>
            {
                try
                {
                    if (platformId is null) new VersionStore(root).CancelPendingActivation();
                    else PlatformHandoff.Cancel(root, platformId);
                    form.Close();
                    shownSlot = null;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { MessageBox.Show(form, T("cancelFailed"), T("title")); }
            };
            actions.Controls.Add(cancel);
        }
        form.FormClosed += (_, _) => { if (ReferenceEquals(notice, form)) notice = null; form.Dispose(); };
        notice = form;
        form.Show(new ExcelOwner());
    }

    private UpdateDiagnostic Record(UpdateStage stage, Exception error)
    {
        var diagnostic = UpdateDiagnostic.Create(stage, error);
        diagnostic.TrySave(root);
        return diagnostic;
    }

    private static void ShowDetails(IWin32Window owner, string language, UpdateDiagnostic diagnostic)
    {
        string T(string key) => UpdateLanguage.Text(language, key);
        using var form = new Form
        {
            Text = T("details"),
            Size = new System.Drawing.Size(620, 420),
            StartPosition = FormStartPosition.CenterParent,
            MinimizeBox = false,
            ShowInTaskbar = false,
            AutoScaleMode = AutoScaleMode.Dpi,
            Padding = new Padding(12)
        };
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(new Label { Text = T("supportHint"), AutoSize = true, MaximumSize = new System.Drawing.Size(560, 0) });
        layout.Controls.Add(new TextBox
        {
            Text = diagnostic.ToJson(),
            ReadOnly = true,
            Multiline = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Dock = DockStyle.Fill,
            AccessibleName = T("details")
        });
        var close = new Button { Text = T("close"), AutoSize = true };
        close.Click += (_, _) => form.Close();
        layout.Controls.Add(close);
        form.Controls.Add(layout);
        form.CancelButton = close;
        form.ShowDialog(owner);
    }

    public void Dispose()
    {
        disposed = true;
        lifetime.Cancel();
        lifetime.Dispose();
        timer.Stop();
        timer.Dispose();
        notice?.Close();
        notice = null;
        if (ReferenceEquals(UpdateCommands.Current, this)) UpdateCommands.Current = null;
    }

    private sealed class ExcelOwner : IWin32Window
    {
        public IntPtr Handle => ExcelDnaUtil.WindowHandle;
    }
}
