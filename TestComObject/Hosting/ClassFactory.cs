namespace TestComObject.Hosting;

[GeneratedComClass]
public partial class ClassFactory : IClassFactory
{
    HRESULT IClassFactory.CreateInstance(nint pUnkOuter, in Guid riid, out nint ppvObject)
    {
        throw new NotImplementedException();
    }

    HRESULT IClassFactory.LockServer(bool fLock)
    {
        return HRESULT.S_OK;
    }
}
