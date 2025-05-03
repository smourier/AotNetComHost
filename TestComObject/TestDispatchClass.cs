namespace TestComObject;

[Guid("be2982f8-a838-497b-88f0-f957f6cf7f87")]
[ProgId("TestComObject.TestDispatchClass")]
[GeneratedComClass]
public partial class TestDispatchClass : IDispatch
{
    HRESULT IDispatch.GetIDsOfNames(in Guid riid, nint[] rgszNames, uint cNames, uint lcid, int[] rgDispId)
    {
        EventProvider.Default.Write("riid:" + riid + " cNames: " + cNames);
        return HRESULT.DISP_E_UNKNOWNNAME;
    }

    HRESULT IDispatch.GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo)
    {
        EventProvider.Default.Write("iTInfo:" + iTInfo + " lcid: " + lcid);
        ppTInfo = 0;
        return HRESULT.E_NOINTERFACE;
    }

    HRESULT IDispatch.GetTypeInfoCount(out uint pctinfo)
    {
        EventProvider.Default.Write();
        pctinfo = 0;
        return HRESULT.S_OK;
    }

    HRESULT IDispatch.Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
    {
        EventProvider.Default.Write("dispIdMember: " + dispIdMember + " riid:" + riid + " v: " + wFlags);
        return HRESULT.DISP_E_MEMBERNOTFOUND;
    }
}
