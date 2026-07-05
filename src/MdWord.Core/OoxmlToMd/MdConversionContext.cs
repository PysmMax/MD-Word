using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Threads the per-conversion state (the main document part for hyperlink/
/// math relationship lookups, resolved style/numbering catalogs, Office's
/// OMML2MML.XSL path, and the warnings sink) through
/// <see cref="DocumentMarkdownBuilder"/>/<see cref="InlineMarkdownBuilder"/>/
/// <see cref="TableMarkdownBuilder"/> — the reverse-direction mirror of
/// <c>MdToOoxml.MathConversionContext</c>.
/// </summary>
internal sealed class MdConversionContext
{
    public MainDocumentPart MainPart { get; }

    public MathXslPaths XslPaths { get; }

    public StyleCatalog Styles { get; }

    public NumberingCatalog Numbering { get; }

    public List<string> Warnings { get; }

    /// <summary>
    /// True when Office's OMML2MML.XSL path is supplied — i.e. formula
    /// conversion should be attempted. False means the documented
    /// "no Office XSLT available" mode: every formula degrades to
    /// <c>`[formula]`</c> literal text plus a warning.
    /// </summary>
    public bool MathActive => XslPaths?.Omml2MmlXsl != null;

    public MdConversionContext(
        MainDocumentPart mainPart,
        MathXslPaths xslPaths,
        StyleCatalog styles,
        NumberingCatalog numbering,
        List<string> warnings)
    {
        MainPart = mainPart;
        XslPaths = xslPaths;
        Styles = styles;
        Numbering = numbering;
        Warnings = warnings;
    }
}
