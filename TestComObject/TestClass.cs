namespace TestComObject;

[Guid("91aae7d8-adc0-4d57-a01c-dfceb85e82e3")]
[GeneratedComClass]
public partial class TestClass : IOleCommandTarget
{
    HRESULT IOleCommandTarget.Exec(in Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, in VARIANT pvaIn, ref VARIANT pvaOut)
    {
        EventProvider.Default.Write("pguidCmdGroup:" + pguidCmdGroup + " nCmdID: " + nCmdID);
        return HRESULT.S_OK;
    }

    HRESULT IOleCommandTarget.QueryStatus(in Guid pguidCmdGroup, uint cCmds, ref OLECMD prgCmds, ref OLECMDTEXT pCmdText)
    {
        EventProvider.Default.Write("pguidCmdGroup:" + pguidCmdGroup + " cCmds: " + cCmds);
        return HRESULT.E_NOTIMPL;
    }
}
