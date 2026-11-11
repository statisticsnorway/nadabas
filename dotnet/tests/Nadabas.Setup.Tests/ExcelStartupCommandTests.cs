using Nadabas.Setup;

namespace Nadabas.Setup.Tests;

public sealed class ExcelStartupCommandTests
{
    [Fact]
    public void MissingPathsDoNotStartExcel()
    {
        var output = new StringWriter();
        Assert.Equal(2, Command.Run(["excel-startup-test"], output, new StringWriter()));
        Assert.Empty(output.ToString());
    }

    [Fact]
    public void RelativePathsDoNotStartExcel()
    {
        var output = new StringWriter();
        Assert.Equal(2, Command.Run(["excel-startup-test", "--source", "NADABAS.XLAM", "--output", "test"], output, new StringWriter()));
        Assert.Empty(output.ToString());
    }

    [Fact]
    public void UnknownFlagDoesNotStartExcel()
    {
        var output = new StringWriter();
        Assert.Equal(2, Command.Run(["excel-startup-test", "--source", Path.Combine(Path.GetTempPath(), "NADABAS.XLAM"),
            "--output", Path.GetTempPath(), "--enable-macros"], output, new StringWriter()));
        Assert.Empty(output.ToString());
    }
}
