using System;
using System.Linq;
using System.Xml.Linq;

namespace MdWord.Core.Math;

/// <summary>
/// LaTeX → OMML conveyor: KaTeX (<see cref="KaTeXRenderer"/>) renders LaTeX
/// to MathML, then <see cref="MathMlToOmmlTransformer"/> runs Office's own
/// MML2OMML.XSL over the extracted <c>&lt;math&gt;</c> element. Never throws —
/// <see cref="TryConvert"/> reports failure so callers can degrade a single
/// formula to literal text without aborting the whole document (see
/// PLAN.md §6 / <see cref="ConvertException"/>'s per-element-failure contract).
/// </summary>
internal static class MathConverter
{
    private static readonly XNamespace MathMlNamespace = "http://www.w3.org/1998/Math/MathML";

    /// <summary>
    /// Converts a single LaTeX formula to the outer XML of an <c>m:oMath</c>
    /// fragment. Returns false (with a human-readable <paramref name="failureReason"/>)
    /// on any failure — invalid LaTeX (KaTeX's <c>throwOnError</c>), a missing/
    /// malformed <c>&lt;math&gt;</c> element in KaTeX's output, or an XSLT
    /// transform failure — instead of throwing.
    /// </summary>
    public static bool TryConvert(string tex, bool displayMode, MathXslPaths xslPaths, out string ommlOuterXml, out string failureReason)
    {
        ommlOuterXml = null;
        failureReason = null;

        try
        {
            var rawMathMl = KaTeXRenderer.RenderToMathMl(tex, displayMode);
            var mathElement = ExtractMathElement(rawMathMl);

            if (mathElement == null)
            {
                failureReason = "KaTeX не повернув елемент <math> з очікуваним простором імен.";
                return false;
            }

            NaryArgumentGrouper.GroupArguments(mathElement);
            var transformed = MathMlToOmmlTransformer.Transform(mathElement, xslPaths.Mml2OmmlXsl);
            ommlOuterXml = NaryHiddenLimitFixup.HideEmptyLimits(transformed);
            return true;
        }
        catch (Exception ex)
        {
            failureReason = ex.Message;
            return false;
        }
    }

    private static XElement ExtractMathElement(string rawMathMlHtml)
    {
        // KaTeX's `output: 'mathml'` wraps the bare <math> in a <span
        // class="katex">...</span> — parse the whole thing as XML (KaTeX's
        // output is well-formed) and pick out the namespaced <math>
        // descendant, rather than string-slicing, so the
        // xmlns="http://www.w3.org/1998/Math/MathML" attribute survives
        // intact. Dropping that namespace is the single most likely way
        // this silently produces an empty m:oMath (MML2OMML.XSL's templates
        // match on the MathML namespace).
        var document = XDocument.Parse(rawMathMlHtml);
        return document.Descendants(MathMlNamespace + "math").FirstOrDefault();
    }
}
