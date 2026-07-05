using System.Linq;
using System.Xml.Linq;

namespace MdWord.Core.Math;

/// <summary>
/// Runs on the OMML string <see cref="MathMlToOmmlTransformer"/> produces,
/// after <see cref="NaryArgumentGrouper"/> has synthesized an <c>msup</c>
/// wrapper (with an empty placeholder superscript) around a bare n-ary
/// operator. MML2OMML.XSL's <c>subHide</c>/<c>supHide</c> are set purely
/// from which element matched (<c>msub</c>/<c>msup</c>/<c>msubsup</c>/
/// <c>munder</c>/<c>mover</c>/<c>munderover</c>) — no combination of those
/// six ever produces "both hidden", which is what an indefinite integral
/// with no limits needs. This patches any <c>m:nary</c> whose
/// <c>&lt;m:sub&gt;</c>/<c>&lt;m:sup&gt;</c> came out empty (only ever true
/// for our synthetic wrapper — a real limit from MathML always has
/// content) to <c>m:val="on"</c>, so Word hides the placeholder instead of
/// showing an empty box next to it.
/// </summary>
internal static class NaryHiddenLimitFixup
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    public static string HideEmptyLimits(string ommlOuterXml)
    {
        var root = XElement.Parse(ommlOuterXml);

        foreach (var nary in root.DescendantsAndSelf(M + "nary"))
        {
            var naryPr = nary.Element(M + "naryPr");
            if (naryPr == null)
            {
                continue;
            }

            HideIfEmpty(nary.Element(M + "sub"), naryPr.Element(M + "subHide"));
            HideIfEmpty(nary.Element(M + "sup"), naryPr.Element(M + "supHide"));
        }

        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static void HideIfEmpty(XElement limit, XElement hideFlag)
    {
        if (limit != null && !limit.HasElements && hideFlag != null)
        {
            hideFlag.SetAttributeValue(M + "val", "on");
        }
    }
}
