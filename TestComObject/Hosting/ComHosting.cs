﻿// the namespace of this ComHosting class is important, it must be <assembly>.Hosting or <assembly
// for the AotNetComHost to find it
namespace TestComObject.Hosting;

[SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification = "Interop code from somewhere else than .NET.")]
public static partial class ComHosting
{
    // TODO: this list represents the COM types you want to expose
    // it's somehow a moral equivalent to setting ComVisible(true) on them
    public static Type[] ComTypes { get; } =
    [
        typeof(TestClass),
        typeof(TestDispatchClass)
    ];

    // create registry entries for all types supported in this module.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllRegisterServer))]
    public static uint DllRegisterServer() => WrapErrors(RegisterServer);

    // remove entries created through DllRegisterServer.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllUnregisterServer))]
    public static uint DllUnregisterServer() => WrapErrors(UnregisterServer);

    // determines whether the module is in use. If not, the caller can unload the DLL from memory.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]
    public static uint DllCanUnloadNow()
    {
        EventProvider.Default.Write($"Path:{DllPath}");
        return HRESULT.S_FALSE;
    }

    // retrieves the class object from a DLL object handler or object application.
    [UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]
    public static unsafe uint DllGetClassObject(nint rclsid, nint riid, nint ppv) => WrapErrors(() =>
    {
        var clsid = *(Guid*)rclsid;
        var iid = *(Guid*)riid;
        var hr = GetClassObject(clsid, iid, out var obj);
        if (hr != HRESULT.S_OK)
            return hr;

        var unk = _wrappers.GetOrCreateComInterfaceForObject(obj!, CreateComInterfaceFlags.None);
        *(nint*)ppv = unk;
        EventProvider.Default.Write($"unk:{unk}");
        return HRESULT.S_OK;
    });

    // handles installation and setup for a module.
    // this one is optional
    [UnmanagedCallersOnly(EntryPoint = nameof(DllInstall))]
    public static uint DllInstall(bool install, nint cmdLinePtr) => WrapErrors(() =>
    {
        EventProvider.Default.Write($"Path:{DllPath}");
        if (cmdLinePtr != 0)
        {
            var cmdLine = Marshal.PtrToStringUni(cmdLinePtr);
            EventProvider.Default.Write($"CmdLine:{cmdLine}");
            if (string.Equals(cmdLine, "user", StringComparison.OrdinalIgnoreCase))
            {
                InstallInHkcu = true;
            }
        }

        return install ? RegisterServer() : UnregisterServer();
    });

    private static HRESULT WrapErrors(Func<HRESULT> action)
    {
        try
        {
            return action();
        }
        catch (SecurityException se)
        {
            // transform this one as a well-known access denied
            EventProvider.Default.Write($"Ex:{se}");
            return HRESULT.E_ACCESSDENIED;
        }
        catch (Exception ex)
        {
            EventProvider.Default.Write($"Ex:{ex}");
            return (uint)ex.HResult;
        }
    }

    private static HRESULT RegisterServer()
    {
        EventProvider.Default.Write($"Path:{DllPath}");
        var root = InstallInHkcu ? Registry.CurrentUser : Registry.LocalMachine;
        foreach (var type in ComTypes)
        {
            RegisterInProcessComObject(root, type, RegistrationDllPath, ThreadingModel);
        }
        return HRESULT.S_OK;
    }

    private static HRESULT UnregisterServer()
    {
        EventProvider.Default.Write($"Path:{DllPath}");
        var root = InstallInHkcu ? Registry.CurrentUser : Registry.LocalMachine;
        foreach (var type in ComTypes)
        {
            UnregisterComObject(root, type);
        }
        return HRESULT.S_OK;
    }

    private static HRESULT GetClassObject(Guid clsid, Guid iid, out object? ppv)
    {
        EventProvider.Default.Write($"Path:{DllPath} CLSID:{clsid} IID:{iid}");
        foreach (var type in ComTypes)
        {
            EventProvider.Default.Write($"Type:{type.FullName} guid:{type.GUID}");
            if (clsid == type.GUID && iid == typeof(IClassFactory).GUID)
            {
                ppv = new ClassFactory(type);
                EventProvider.Default.Write($"ppv:{ppv}");
                return HRESULT.S_OK;
            }
        }
        ppv = null;
        EventProvider.Default.Write($"E_NOINTERFACE");
        return HRESULT.E_NOINTERFACE;
    }

    public static bool InstallInHkcu { get; set; } = false;
    public static string? ThreadingModel { get; set; } // default is Both
    public static string DllPath { get; } = BuildPath();
    public static string RegistrationDllPath
#if DEBUG
        => _thunkDllPath ?? DllPath;
#else
        => DllPath;
#endif

    internal static readonly StrategyBasedComWrappers _wrappers = new();
    private const string ClassesRegistryKey = @"Software\Classes";
    private const string ClsidRegistryKey = ClassesRegistryKey + @"\CLSID";

    public static void RegisterInProcessComObject(RegistryKey root, Type type, string assemblyPath, string? threadingModel = null)
    {
        ArgumentNullException.ThrowIfNull(root);
        ArgumentNullException.ThrowIfNull(type);
        ArgumentNullException.ThrowIfNull(assemblyPath);

        threadingModel = threadingModel?.Trim() ?? "Both";
        EventProvider.Default.Write($"Registering {type.FullName} from {assemblyPath} with threading model '{threadingModel}'...");
        using var serverKey = EnsureWritableSubKey(root, Path.Combine(ClsidRegistryKey, type.GUID.ToString("B"), "InprocServer32"));
        serverKey.SetValue(null, assemblyPath);
        serverKey.SetValue("ThreadingModel", threadingModel);

        // ProgId is optional, make sure BuiltInComInteropSupport property is set to true in csproj to avoid it to be trimmed out during AOT publish
        var att = type.GetCustomAttribute<ProgIdAttribute>();
        if (att != null && !string.IsNullOrWhiteSpace(att.Value))
        {
            var progid = att.Value.Trim();
            using var key = EnsureWritableSubKey(root, Path.Combine(ClsidRegistryKey, type.GUID.ToString("B")));
            using var progIdKey = EnsureWritableSubKey(key, "ProgId");
            progIdKey.SetValue(null, progid);

            using var ckey = EnsureWritableSubKey(root, Path.Combine(ClassesRegistryKey, progid, "CLSID"));
            ckey.SetValue(null, type.GUID.ToString("B"));
        }
        EventProvider.Default.Write($"Registered {type.FullName}.");
    }

    public static void UnregisterComObject(RegistryKey root, Type type)
    {
        ArgumentNullException.ThrowIfNull(root);
        ArgumentNullException.ThrowIfNull(type);

        EventProvider.Default.Write($"Unregistering {type.FullName}...");
        using var key = root.OpenSubKey(ClsidRegistryKey, true);
        key?.DeleteSubKeyTree(type.GUID.ToString("B"), false);

        // ProgId is optional, make sure BuiltInComInteropSupport property is set to true in csproj to avoid it to be trimmed out during AOT publish
        var att = type.GetCustomAttribute<ProgIdAttribute>();
        if (att != null && !string.IsNullOrWhiteSpace(att.Value))
        {
            var progid = att.Value.Trim();
            using var ckey = root.OpenSubKey(ClassesRegistryKey, true);
            ckey?.DeleteSubKeyTree(progid, false);
        }

        EventProvider.Default.Write($"Unregistered {type.FullName}.");
    }

    public static RegistryKey EnsureWritableSubKey(RegistryKey root, string name)
    {
        ArgumentNullException.ThrowIfNull(root);
        ArgumentNullException.ThrowIfNull(name);

        var key = root.OpenSubKey(name, true);
        if (key != null)
            return key;

        var parentName = Path.GetDirectoryName(name);
        if (string.IsNullOrEmpty(parentName))
            return root.CreateSubKey(name);

        using var parentKey = EnsureWritableSubKey(root, parentName);
        return parentKey.CreateSubKey(Path.GetFileName(name));
    }

#if DEBUG
    // this only works if not published as AOT
    [UnconditionalSuppressMessage("SingleFile", "IL3000:Avoid accessing Assembly file path when publishing as a single file", Justification = "Only used when not published as AOT")]
    private static string BuildPath() => typeof(ComHosting).Assembly.Location;

    private static string? _thunkDllPath = DllPath;

    [UnmanagedCallersOnly(EntryPoint = nameof(DllThunkInit))]
    public unsafe static uint DllThunkInit(nint thunkDllPathPtr)
    {
        var types = ComTypes;
        _thunkDllPath = Marshal.PtrToStringUni(thunkDllPathPtr);
        EventProvider.Default.Write($"Path:{DllPath} ThunkDllPathPtr:{_thunkDllPath}");
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
            throw new Win32Exception(Marshal.GetLastPInvokeError());

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
