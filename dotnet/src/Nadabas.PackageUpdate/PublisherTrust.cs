using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;

[assembly: DefaultDllImportSearchPaths(DllImportSearchPath.System32)]

namespace Nadabas.PackageUpdate;

internal static class PublisherTrust
{
    internal const string Thumbprint = "A70F7882B2F9F5C6D23A22C8AEC70211854B9EAF";

    // Windows Authenticode policy, not merely extraction of a certificate.
    // https://learn.microsoft.com/windows/win32/api/wintrust/nf-wintrust-winverifytrust
    internal static void Verify(string path)
    {
        if (!OperatingSystem.IsWindows()) throw new IOException("Windows signature verification is required.");
        var name = Marshal.StringToCoTaskMemUni(path);
        var info = new FileInfoNative { Size = (uint)Marshal.SizeOf<FileInfoNative>(), Path = name };
        var pointer = Marshal.AllocHGlobal(Marshal.SizeOf<FileInfoNative>());
        Marshal.StructureToPtr(info, pointer, false);
        var data = new TrustData
        {
            Size = (uint)Marshal.SizeOf<TrustData>(),
            Ui = 2,
            Revocation = 1,
            Choice = 1,
            File = pointer,
            Action = 1
        };
        var action = new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");
        try
        {
            var result = WinVerifyTrust(new IntPtr(-1), ref action, ref data);
            if (result != 0) throw new IOException($"Authenticode validation failed: 0x{result:X8}.");
            // Pin the signer from the verified trust state, not a separately
            // parsed certificate that might belong to a different signature.
            var provider = WTHelperProvDataFromStateData(data.State);
            var signer = provider == IntPtr.Zero ? IntPtr.Zero : WTHelperGetProvSignerFromChain(provider, 0, false, 0);
            var providerCert = signer == IntPtr.Zero ? IntPtr.Zero : WTHelperGetProvCertFromChain(signer, 0);
            if (providerCert == IntPtr.Zero) throw new IOException("Verified signer certificate is unavailable.");
            var certificatePointer = Marshal.PtrToStructure<ProviderCertificatePrefix>(providerCert).Certificate;
            if (certificatePointer == IntPtr.Zero) throw new IOException("Verified certificate context is unavailable.");
            var context = Marshal.PtrToStructure<CertificateContext>(certificatePointer);
            if (context.Length is 0 or > 1048576 || context.Encoded == IntPtr.Zero) throw new IOException("Invalid signer certificate size.");
            var encoded = new byte[context.Length];
            Marshal.Copy(context.Encoded, encoded, 0, encoded.Length);
            using var certificate = X509CertificateLoader.LoadCertificate(encoded);
            if (!string.Equals(certificate.GetCertHashString(), Thumbprint, StringComparison.OrdinalIgnoreCase))
                throw new IOException("Package is not signed by the approved NADABAS publisher.");
        }
        finally
        {
            data.Action = 2;
            WinVerifyTrust(new IntPtr(-1), ref action, ref data);
            Marshal.FreeHGlobal(pointer);
            Marshal.FreeCoTaskMem(name);
        }
    }

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern int WinVerifyTrust(IntPtr window, ref Guid action, ref TrustData data);

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern IntPtr WTHelperProvDataFromStateData(IntPtr state);
    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern IntPtr WTHelperGetProvSignerFromChain(IntPtr provider, uint signer, [MarshalAs(UnmanagedType.Bool)] bool counterSigner, uint counter);
    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern IntPtr WTHelperGetProvCertFromChain(IntPtr signer, uint index);

    [StructLayout(LayoutKind.Sequential)]
    private struct ProviderCertificatePrefix { public uint Size; public IntPtr Certificate; }
    [StructLayout(LayoutKind.Sequential)]
    private struct CertificateContext
    {
        public uint Encoding;
        public IntPtr Encoded;
        public uint Length;
        public IntPtr Info, Store;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FileInfoNative { public uint Size; public IntPtr Path, Handle, Subject; }

    [StructLayout(LayoutKind.Sequential)]
    private struct TrustData
    {
        public uint Size;
        public IntPtr Policy, Client;
        public uint Ui, Revocation, Choice;
        public IntPtr File;
        public uint Action;
        public IntPtr State, Url;
        public uint Flags, Context;
        public IntPtr Settings;
    }
}
