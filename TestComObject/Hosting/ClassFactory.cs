namespace TestComObject.Hosting;

[GeneratedComClass]
internal partial class ClassFactory : IClassFactory
{
    public ClassFactory(Type type)
    {
        ArgumentNullException.ThrowIfNull(type);
        Type = type;
    }

    public Type Type { get; }

    HRESULT IClassFactory.CreateInstance(nint pUnkOuter, in Guid riid, out nint ppvObject)
    {
        EventProvider.Default.Write($"pUnkOuter:{pUnkOuter} riid:{riid}");
        ppvObject = 0;
        var obj = Activator.CreateInstance(Type, true);
        if (obj is null)
            return HRESULT.E_NOINTERFACE;

        var unk = ComHosting._wrappers.GetOrCreateComInterfaceForObject(obj, CreateComInterfaceFlags.None);
        if (riid != typeof(IUnknown).GUID)
        {
            var hr = Marshal.QueryInterface(unk, riid, out ppvObject);
            Marshal.Release(unk);
            if (hr < 0)
                return (uint)hr;
        }
        else
        {
            ppvObject = unk;
        }
        return HRESULT.S_OK;
    }

    HRESULT IClassFactory.LockServer(bool fLock)
    {
        EventProvider.Default.Write($"lock:{fLock}");
        return HRESULT.S_OK;
    }
}
