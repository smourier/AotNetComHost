namespace TestComObject.Interop;

// https://learn.microsoft.com/office/client-developer/outlook/mapi/hresult
public partial struct HRESULT(uint value) : IEquatable<HRESULT>
{
    public static readonly HRESULT S_OK = new();
    public static readonly HRESULT S_FALSE = new(0x00000001);
    public static readonly HRESULT E_NOINTERFACE = new(0x80004002);
    public static readonly HRESULT E_INVALIDARG = new(0x80070057);
    public static readonly HRESULT E_FAIL = new(0x80004005);
    public static readonly HRESULT E_NOTIMPL = new(0x80004001);
    public static readonly HRESULT E_ACCESSDENIED = new(0x80070005);
    public static readonly HRESULT DISP_E_UNKNOWNNAME = new(0x80020006);
    public static readonly HRESULT DISP_E_MEMBERNOTFOUND = new(0x80020003);
    public static readonly HRESULT CLASS_E_CLASSNOTAVAILABLE = new(0x80040111);

    public uint Value = value;

    public override readonly bool Equals(object? obj) => obj is HRESULT value && Equals(value);
    public readonly bool Equals(HRESULT other) => other.Value == Value;
    public override readonly int GetHashCode() => Value.GetHashCode();
    public static bool operator ==(HRESULT left, HRESULT right) => left.Equals(right);
    public static bool operator !=(HRESULT left, HRESULT right) => !left.Equals(right);
    public static implicit operator uint(HRESULT value) => value.Value;
    public static implicit operator HRESULT(uint value) => new(value);
}
