using Nadabas.PackageUpdate;
using System.Text.Json;

namespace Nadabas.Setup.Tests;

public sealed class UpdateDiagnosticTests
{
    [Theory]
    [InlineData(404)]
    [InlineData(403)]
    [InlineData(503)]
    public void HttpFailureKeepsStatusWithoutRequestOrResponseContent(int status)
    {
        var error = new HttpRequestException("https://private/?token=secret", null,
            (System.Net.HttpStatusCode)status);
        var report = UpdateDiagnostic.Create(UpdateStage.Download, new IOException("private", error));
        Assert.Null(report.Failures[0].HttpStatusCode);
        Assert.Equal(status, report.Failures[1].HttpStatusCode);
        Assert.DoesNotContain("private", report.ToJson());
        Assert.DoesNotContain("secret", report.ToJson());
    }

    [Fact]
    public void NetworkFailureAndOlderReportsHaveNoInventedHttpStatus()
    {
        var report = UpdateDiagnostic.Create(UpdateStage.Download, new HttpRequestException("private"));
        Assert.Null(report.Failures[0].HttpStatusCode);
        var legacy = JsonSerializer.Deserialize<UpdateFailure>("{\"Type\":\"HttpRequestException\",\"HResult\":\"0x80131500\"}");
        Assert.Null(legacy!.HttpStatusCode);
    }

    [Fact]
    public void ReportKeepsNestedErrorCodesWithoutSensitiveExceptionContent()
    {
        var inner = new FileNotFoundException("password=secret", "private-workbook.xlsx");
        inner.Data["token"] = "private-token";
        var report = UpdateDiagnostic.Create(UpdateStage.Verification, new IOException("https://private/?key=secret", inner));
        var json = report.ToJson();
        Assert.Equal(2, report.Failures.Count);
        Assert.Equal("FileNotFoundException", report.Failures[1].Type);
        Assert.Equal($"0x{inner.HResult:X8}", report.Failures[1].HResult);
        Assert.DoesNotContain("secret", json);
        Assert.DoesNotContain("private", json);
        Assert.Equal(report.Reference, JsonSerializer.Deserialize<UpdateDiagnostic>(json)!.Reference);
    }

    [Fact]
    public void RepeatedFailuresReplaceLastReportAndLeaveNoTemporaryFiles()
    {
        var root = Path.Combine(Path.GetTempPath(), "nadabas-diagnostic-" + Guid.NewGuid().ToString("N"));
        try
        {
            var first = UpdateDiagnostic.Create(UpdateStage.Download, new IOException());
            var second = UpdateDiagnostic.Create(UpdateStage.Download, new UnauthorizedAccessException());
            Assert.NotNull(first.TrySave(root));
            var path = second.TrySave(root);
            Assert.NotNull(path);
            Assert.Equal(second.Reference, JsonSerializer.Deserialize<UpdateDiagnostic>(File.ReadAllText(path))!.Reference);
            Assert.Single(Directory.GetFiles(Path.Combine(root, "diagnostics")));
        }
        finally { if (Directory.Exists(root)) Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public void StorageFailureStillLeavesReportAvailableForDetails()
    {
        var file = Path.GetTempFileName();
        try
        {
            var report = UpdateDiagnostic.Create(UpdateStage.Restart, new IOException());
            Assert.Null(report.TrySave(file));
            Assert.Contains("Restart", report.ToJson());
            Assert.Null((report with { Stage = "../other" }).TrySave(Path.GetTempPath()));
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void ExceptionDepthIsBounded()
    {
        Exception error = new IOException();
        for (var i = 0; i < 30; i++) error = new IOException("sensitive", error);
        Assert.Equal(8, UpdateDiagnostic.Create(UpdateStage.Verification, error).Failures.Count);
    }
}
