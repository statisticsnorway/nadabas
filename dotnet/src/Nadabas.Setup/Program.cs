namespace Nadabas.Setup;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args) => Command.Run(args, Console.Out, Console.Error);
}
