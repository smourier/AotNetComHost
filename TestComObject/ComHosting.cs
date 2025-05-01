namespace TestComObject;

public static class ComHosting
{
    public delegate uint DllRegisterServerDelegate();

    [UnmanagedCallersOnly(EntryPoint = nameof(DllRegisterServer))]
    public static uint DllRegisterServer()
    {
        EventProvider.Default.Write();
        return 0;
    }

    [UnmanagedCallersOnly(EntryPoint = nameof(DllUnregisterServer))]
    public static uint DllUnregisterServer()
    {
        EventProvider.Default.Write();
        return 0;
    }
}
