namespace TestComObject;

public static class ComHosting
{
    private const uint E_NOINTERFACE = 0x80004002;
    private const uint E_FAIL = 0x80004005;
    private const uint E_NOTIMPL = 0x80004001;

    [UnmanagedCallersOnly(EntryPoint = nameof(DllRegisterServer))]
    public static uint DllRegisterServer()
    {
        EventProvider.Default.Write();
        return E_NOTIMPL;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(DllUnregisterServer))]
    public static uint DllUnregisterServer()
    {
        EventProvider.Default.Write();
        return E_NOTIMPL;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]
    public static uint DllCanUnloadNow()
    {
        EventProvider.Default.Write();
        return E_NOTIMPL;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(DllInstall))]
    public static uint DllGetClassObject(nint rclsid, nint riid, nint ppv)
    {
        EventProvider.Default.Write();
        return E_NOTIMPL;
    }

    // this one is optional
    [UnmanagedCallersOnly(EntryPoint = nameof(DllInstall))]
    public static uint DllInstall(bool install, nint cmdLine)
    {
        EventProvider.Default.Write();
        return E_NOTIMPL;
    }
}
