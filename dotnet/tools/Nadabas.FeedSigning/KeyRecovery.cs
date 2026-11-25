using Chaos.NaCl;
using System.Security.Cryptography;

namespace Nadabas.FeedSigning;

internal static class KeyRecovery
{
    internal static void Run(string command, string input, string output, string? expectedPublicKey)
    {
        if (!OperatingSystem.IsWindows()) throw new IOException("Windows user protection is required.");
        if (!Path.IsPathFullyQualified(input) || !Path.IsPathFullyQualified(output))
            throw new IOException("Use absolute paths.");
        if (File.Exists(output) || Directory.Exists(output)) throw new IOException("Choose a new output file.");
        if (new FileInfo(input).Length > 4096) throw new IOException("Input file is too large.");
        var password = ReadPassword();
        byte[]? seed = null;
        byte[]? privateKey = null;
        try
        {
            byte[] result;
            if (command == "backup-key")
            {
                var confirmation = ReadPassword();
                try
                {
                    if (!password.AsSpan().SequenceEqual(confirmation)) throw new IOException("Passphrases differ.");
                }
                finally { Array.Clear(confirmation); }
                seed = ProtectedData.Unprotect(File.ReadAllBytes(input), null, DataProtectionScope.CurrentUser);
                result = PortableSeedBackup.Encrypt(seed, password);
            }
            else
            {
                seed = PortableSeedBackup.Decrypt(File.ReadAllBytes(input), password);
                Ed25519.KeyPairFromSeed(out var publicKey, out privateKey, seed);
                if (Convert.ToBase64String(publicKey) != expectedPublicKey)
                    throw new IOException("Recovered key differs from the independently recorded public key.");
                result = ProtectedData.Protect(seed, null, DataProtectionScope.CurrentUser);
            }
            using var file = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            file.Write(result);
            file.Flush(true);
            Console.WriteLine(command == "backup-key" ? "Encrypted recovery backup written." : "Key recovered and protected for this Windows user.");
        }
        finally
        {
            Array.Clear(password);
            if (seed is not null) CryptographicOperations.ZeroMemory(seed);
            if (privateKey is not null) CryptographicOperations.ZeroMemory(privateKey);
        }
    }

    private static char[] ReadPassword()
    {
        if (Console.IsInputRedirected) throw new IOException("Enter the recovery passphrase in an interactive terminal.");
        Console.Write("Recovery passphrase (hidden): ");
        var buffer = new char[1024];
        var length = 0;
        try
        {
            while (true)
            {
                var key = Console.ReadKey(intercept: true);
                if (key.Key == ConsoleKey.Enter) break;
                if (key.Key == ConsoleKey.Backspace)
                {
                    if (length > 0) buffer[--length] = '\0';
                }
                else if (!char.IsControl(key.KeyChar))
                {
                    if (length == buffer.Length) throw new IOException("Passphrase is too long.");
                    buffer[length++] = key.KeyChar;
                }
            }
            Console.WriteLine();
            return buffer[..length];
        }
        finally { Array.Clear(buffer); }
    }
}
