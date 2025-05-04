namespace TestComObject.Interop;

[GeneratedComInterface, Guid("00000001-0000-0000-c000-000000000046")]
public partial interface IClassFactory
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT CreateInstance(nint pUnkOuter, in Guid riid, out nint /* void */ ppvObject);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock);
}
