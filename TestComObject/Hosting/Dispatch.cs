namespace TestComObject.Hosting;

// don't add GeneratedComClass, it will break with AOT publish for some reason
#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable SYSLIB1097 // Add 'GeneratedComClassAttribute' to enable passing objects of this type to COM
public abstract partial class Dispatch<T>
    : IDispatch where T : new() // this is needed to ensure derived class' parameterless constructor will not be trimmed out with AOT publish
#pragma warning restore SYSLIB1097 // Add 'GeneratedComClassAttribute' to enable passing objects of this type to COM
#pragma warning restore IDE0079 // Remove unnecessary suppression
{
    HRESULT IDispatch.GetIDsOfNames(in Guid riid, nint[] rgszNames, uint cNames, uint lcid, int[] rgDispId)
    {
        EventProvider.Default.Write($"riid:{riid} cNames:{cNames}");
        if (rgszNames == null || rgszNames.Length == 0)
            return HRESULT.E_INVALIDARG;

        for (var i = 0; i < cNames; i++)
        {
            var name = Marshal.PtrToStringUni(rgszNames[i]);
            if (string.IsNullOrWhiteSpace(name))
                return HRESULT.E_INVALIDARG;

            var dispid = GetDispId(name);
            rgDispId[i] = dispid;
            EventProvider.Default.Write($"name[{i}]:'{name}' => {rgDispId[i]}");
        }

        return HRESULT.S_OK;
    }

    HRESULT IDispatch.GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo)
    {
        EventProvider.Default.Write($"iTInfo:{iTInfo} lcid:{lcid}");
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
        return Invoke(dispIdMember, wFlags, pDispParams, pVarResult);
    }

    protected abstract int GetDispId(string name); // must be > 0 for a known member
    protected abstract HRESULT Invoke(int dispId, DISPATCH_FLAGS flags, DISPPARAMS parameters, nint pVarResult);
}
