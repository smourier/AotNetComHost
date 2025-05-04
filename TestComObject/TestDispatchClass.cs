namespace TestComObject;

[Guid("be2982f8-a838-497b-88f0-f957f6cf7f87")]
[ProgId("TestComObject.TestDispatchClass")]
[GeneratedComClass]
public partial class TestDispatchClass : Dispatch<TestDispatchClass>
{
    public double ComputePi()
    {
        Console.WriteLine("computed!");
        return Math.PI;
    }

    protected override int GetDispId(string name)
    {
        if (string.Compare(name, "ComputePi", StringComparison.OrdinalIgnoreCase) == 0)
            return 1;

        return 0;
    }

    protected override unsafe HRESULT Invoke(int dispId, DISPATCH_FLAGS flags, DISPPARAMS parameters, nint pVarResult)
    {
        if (dispId == 1)
        {
            var value = ComputePi();
            var pv = (VARIANT*)pVarResult;
            pv->Anonymous.Anonymous.vt = VARENUM.VT_R8;
            pv->Anonymous.Anonymous.Anonymous.dblVal = value;
            return HRESULT.S_OK;
        }
        return HRESULT.DISP_E_MEMBERNOTFOUND;
    }
}
