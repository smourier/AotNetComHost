namespace TestComObject.Interop;

internal partial struct OLECMDTEXT
{
    public uint cmdtextf;
    public uint cwActual;
    public uint cwBuf;
    public char rgwz; // variable-length array placeholder
}
