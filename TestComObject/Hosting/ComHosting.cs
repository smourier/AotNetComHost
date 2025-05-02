namespace TestComObject.Hosting;

[SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Interop code from somewhere else than .NET.")]
public static partial class ComHosting
{
    public static Type[] ComTypes { get; } =
    [
        typeof(TestClass),
        typeof(TestDispatchClass)
    ];

    private const uint E_NOINTERFACE = 0x80004002;
    private const uint E_FAIL = 0x80004005;
    private const uint E_NOTIMPL = 0x80004001;

    [UnmanagedCallersOnly(EntryPoint = nameof(DllRegisterServer))]
    public static uint DllRegisterServer()
    {
        foreach (var type in ComTypes)
        {
            EventProvider.Default.Write("type:" + type.FullName);
        }
        EventProvider.Default.Write("Path:" + DllPath);
        return E_NOTIMPL;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(DllUnregisterServer))]
    public static uint DllUnregisterServer()
    {
        EventProvider.Default.Write("Path:" + DllPath);
        return E_NOTIMPL;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]
    public static uint DllCanUnloadNow()
    {
        EventProvider.Default.Write("Path:" + DllPath);
        return E_NOTIMPL;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]
    public static uint DllGetClassObject(nint rclsid, nint riid, nint ppv)
    {
        EventProvider.Default.Write("Path:" + DllPath);
        return E_NOTIMPL;
    }

    // this one is optional
    [UnmanagedCallersOnly(EntryPoint = nameof(DllInstall))]
    public static uint DllInstall(bool install, nint cmdLine)
    {
        EventProvider.Default.Write("Path:" + DllPath);
        return E_NOTIMPL;
    }

    public static string DllPath { get; } = BuildPath();

#if DEBUG
    // this only works if not published as AOT
    [UnconditionalSuppressMessage("SingleFile", "IL3000:Avoid accessing Assembly file path when publishing as a single file", Justification = "<Pending>")]
    private static string BuildPath() => typeof(ComHosting).Assembly.Location;

    private static string? _thunkDllPath = DllPath;

    [UnmanagedCallersOnly(EntryPoint = nameof(DllThunkInit))]
    public unsafe static uint DllThunkInit(nint thunkDllPathPtr)
    {
        var types = ComTypes;
        _thunkDllPath = Marshal.PtrToStringUni(thunkDllPathPtr);
        EventProvider.Default.Write("Path:" + DllPath + " ThunkDllPathPtr: " + _thunkDllPath);
        return 0;
    }
#else
    // this only works if published as AOT
    private static unsafe string BuildPath()
    {
        var ptr = GetModuleFunctionPointer();
        const int GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS = 4;
        const int GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT = 2;
        if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, (nint)ptr, out var module))
            throw new Win32Exception(Marshal.GetLastPInvokeError());

        var moduleName = stackalloc char[260];
        if (GetModuleFileNameW(module, moduleName, 260) == 0)
            throw new Win32Exception(Marshal.GetLastWin32Error());

        return new string(moduleName);
    }

    private static unsafe delegate* unmanaged<uint> GetModuleFunctionPointer() => &DllUnregisterServer; // any func pointer will do

    [LibraryImport("kernel32", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool GetModuleHandleExW(uint dwFlags, nint lpModuleName, out nint phModule);

    [LibraryImport("kernel32", SetLastError = true)]
    private unsafe static partial int GetModuleFileNameW(nint hModule, char* lpFilename, uint nSize);
#endif
}
