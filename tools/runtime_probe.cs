// Engineering probe for the native hosting sequence used by Excel-DNA 1.9.
// Run in Windows PowerShell (Framework), not a process already hosting CoreCLR.
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

public static class NadabasRuntimeProbe
{
    [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibraryW(string name);
    [DllImport("kernel32", CharSet = CharSet.Ansi)]
    private static extern IntPtr GetProcAddress(IntPtr library, string name);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
    private delegate int FindHost(StringBuilder buffer, ref UIntPtr size, IntPtr parameters);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
    private delegate int Initialize(string config, IntPtr parameters, out IntPtr context);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
    private delegate int Property(IntPtr context, string name, out IntPtr value);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    private delegate int Close(IntPtr context);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    private delegate int RuntimeDelegate(IntPtr context, int kind, out IntPtr value);

    private static Delegate Export(IntPtr library, string name, Type type)
    {
        var pointer = GetProcAddress(library, name);
        if (pointer == IntPtr.Zero) throw new InvalidOperationException("Missing native export: " + name);
        return Marshal.GetDelegateForFunctionPointer(pointer, type);
    }

    public static string Run(string nativeHost, string root, string config)
    {
        var library = LoadLibraryW(nativeHost);
        if (library == IntPtr.Zero) throw new InvalidOperationException("Cannot load probe nethost");
        var find = (FindHost)Export(library, "get_hostfxr_path", typeof(FindHost));
        var buffer = new StringBuilder(32768);
        var size = new UIntPtr(32768);
        var result = find(buffer, ref size, IntPtr.Zero);
        if (result != 0) throw new InvalidOperationException("Host discovery failed: " + result);
        var prefix = Path.GetFullPath(root).TrimEnd('\\') + "\\";
        var host = Path.GetFullPath(buffer.ToString());
        if (!host.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Host discovery escaped the local runtime: " + host);
        library = LoadLibraryW(host);
        if (library == IntPtr.Zero) throw new InvalidOperationException("Cannot load local hostfxr");
        var initialize = (Initialize)Export(library, "hostfxr_initialize_for_runtime_config", typeof(Initialize));
        var property = (Property)Export(library, "hostfxr_get_runtime_property_value", typeof(Property));
        var close = (Close)Export(library, "hostfxr_close", typeof(Close));
        IntPtr context;
        result = initialize(config, IntPtr.Zero, out context);
        if (result < 0) throw new InvalidOperationException("Local Desktop Runtime initialization failed: " + result);
        try
        {
            IntPtr value;
            if (property(context, "FX_DEPS_FILE", out value) != 0)
                throw new InvalidOperationException("Runtime dependency path unavailable");
            var dependency = Marshal.PtrToStringUni(value);
            if (!Path.GetFullPath(dependency).StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Runtime resolution escaped local directory");
            var getDelegate = (RuntimeDelegate)Export(library, "hostfxr_get_runtime_delegate", typeof(RuntimeDelegate));
            IntPtr loader;
            // hdt_load_assembly_and_get_function_pointer, also used by Excel-DNA.
            if (getDelegate(context, 5, out loader) < 0 || loader == IntPtr.Zero)
                throw new InvalidOperationException("Local runtime activation failed");
            return dependency;
        }
        finally { close(context); }
    }
}
