using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Resolves a paragraph's Markdown heading level (1–6) via the same
/// three-way fallback the Phase 4 brief mandates, mirroring
/// <c>MdToOoxml.StylesPartBuilder</c>'s <c>HeadingN</c> convention in
/// reverse:
/// <list type="number">
/// <item>styleId matches <c>Heading[1-9]</c> (case-insensitive) — our own
/// generated documents;</item>
/// <item>otherwise, the style's declared <c>w:name</c> is <c>"heading N"</c>
/// (case-insensitive) — catches a foreign/localized document where someone
/// renamed or duplicated the style but Word's invariant name survived;</item>
/// <item>otherwise, <c>w:outlineLvl</c> is present directly on the
/// paragraph (0–5) — catches a paragraph with outline level set but no
/// heading style at all.</item>
/// </list>
/// Word supports up to <c>Heading9</c>/outline level 8, but Markdown only
/// has H1–H6 — levels beyond 6 clamp to 6 rather than being dropped.
/// </summary>
internal static class HeadingResolver
{
    private static readonly Regex HeadingStyleIdPattern = new(@"^Heading\s*([1-9])$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex HeadingNamePattern = new(@"^heading\s*([1-9])$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static int? ResolveLevel(Paragraph paragraph, StyleCatalog styleCatalog)
    {
        var paragraphProperties = paragraph.ParagraphProperties;
        var styleId = paragraphProperties?.ParagraphStyleId?.Val?.Value;

        if (styleId != null)
        {
            var idMatch = HeadingStyleIdPattern.Match(styleId);
            if (idMatch.Success)
            {
                return ClampLevel(int.Parse(idMatch.Groups[1].Value));
            }

            var styleName = styleCatalog?.GetStyleName(styleId);
            if (styleName != null)
            {
                var nameMatch = HeadingNamePattern.Match(styleName.Trim());
                if (nameMatch.Success)
                {
                    return ClampLevel(int.Parse(nameMatch.Groups[1].Value));
                }
            }
        }

        var outlineLevel = paragraphProperties?.OutlineLevel?.Val?.Value;
        if (outlineLevel.HasValue && outlineLevel.Value >= 0 && outlineLevel.Value <= 5)
        {
            return outlineLevel.Value + 1;
        }

        return null;
    }

    private static int ClampLevel(int level) => level switch
    {
        < 1 => 1,
        > 6 => 6,
        _ => level,
    };
}
