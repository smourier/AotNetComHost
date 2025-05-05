namespace TestComObject.Interop;

public struct OLECMDTEXT
{
    public uint cmdtextf;
    public uint cwActual;
    public uint cwBuf;
    public char rgwz; // variable-length array placeholder
}
