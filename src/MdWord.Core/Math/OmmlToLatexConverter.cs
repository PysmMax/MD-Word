using System;
using System.Xml.Linq;

namespace MdWord.Core.Math;

/// <summary>
/// OMML → LaTeX conveyor: <see cref="OmmlToMathMlTransformer"/> runs Office's
/// own OMML2MML.XSL over a single <c>m:oMath</c> element, then
/// <see cref="MathMlToLatexTransformer"/> runs the vendored xsltml library
/// over the resulting MathML. Mirror of <see cref="MathConverter"/> (the
/// forward LaTeX → OMML conveyor): never throws — <see cref="TryConvert"/>
/// reports failure so callers can degrade a single formula to literal
/// <c>`[formula]`</c> text without aborting the whole document (Phase 4
/// brief: "a formula that fails this pipeline entirely must degrade to
/// literal `[formula]` text + a Warning, never throw").
/// </summary>
internal static class OmmlToLatexConverter
{
    /// <summary>
    /// Converts a single OMML <c>m:oMath</c> element's outer XML to LaTeX
    /// (no surrounding <c>$...$</c>/<c>$$...$$</c> delimiters — the caller
    /// in <c>MdWord.Core.OoxmlToMd</c> applies those based on whether the
    /// source was an inline <c>m:oMath</c> or a block <c>m:oMathPara</c>).
    /// Returns false (with a human-readable <paramref name="failureReason"/>)
    /// on any failure — malformed OMML, a missing/incomplete
    /// <see cref="MathXslPaths"/>, or an XSLT transform failure.
    /// </summary>
    public static bool TryConvert(string ommlOuterXml, MathXslPaths xslPaths, out string latex, out string failureReason)
    {
        latex = null;
        failureReason = null;

        if (xslPaths?.Omml2MmlXsl == null)
        {
            failureReason = "Omml2MmlXsl not provided — formula-to-LaTeX conversion is unavailable.";
            return false;
        }

        try
        {
            var ommlElement = XElement.Parse(ommlOuterXml);
            var mathMlElement = OmmlToMathMlTransformer.Transform(ommlElement, xslPaths.Omml2MmlXsl);
            var rawLatex = MathMlToLatexTransformer.Transform(mathMlElement);
            var trimmed = rawLatex?.Trim();

            if (string.IsNullOrEmpty(trimmed))
            {
                failureReason = "mmltex returned empty LaTeX.";
                return false;
            }

            latex = trimmed;
            return true;
        }
        catch (Exception ex)
        {
            failureReason = ex.Message;
            return false;
        }
    }
}
