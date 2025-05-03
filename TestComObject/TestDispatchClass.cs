namespace TestComObject;

[Guid("be2982f8-a838-497b-88f0-f957f6cf7f87")]
[ProgId("TestComObject.TestDispatchClass")]
[GeneratedComClass]
public partial class TestDispatchClass : IDispatch
{
    HRESULT IDispatch.GetIDsOfNames(in Guid riid, nint[] rgszNames, uint cNames, uint lcid, int[] rgDispId)
    {
        throw new NotImplementedException();
    }

    HRESULT IDispatch.GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo)
    {
        throw new NotImplementedException();
    }

    HRESULT IDispatch.GetTypeInfoCount(out uint pctinfo)
    {
        throw new NotImplementedException();
    }

    HRESULT IDispatch.Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
    {
        throw new NotImplementedException();
    }
}
