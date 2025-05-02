namespace TestComObject.Interop;

[GeneratedComInterface, Guid("b722bccb-4e68-101b-a2bc-00aa00404770")]
internal partial interface IOleCommandTarget
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT QueryStatus(in Guid pguidCmdGroup, uint cCmds, ref OLECMD prgCmds, ref OLECMDTEXT pCmdText);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Exec(in Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, in VARIANT pvaIn, ref VARIANT pvaOut);
}
