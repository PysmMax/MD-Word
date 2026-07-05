using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Looks up a document's own <c>styles.xml</c> (<see cref="Styles"/>) by
/// styleId, so <see cref="HeadingResolver"/> can check a style's declared
/// <c>w:name</c> value even when the styleId itself isn't literally
/// <c>HeadingN</c> — the "chuzhyi document with renamed/localized heading
/// styles" case the Phase 4 brief calls out. Read-only mirror of
/// <c>MdToOoxml.StylesPartBuilder</c>, which only ever writes.
/// </summary>
internal sealed class StyleCatalog
{
    private readonly Dictionary<string, Style> _stylesById = new(StringComparer.OrdinalIgnoreCase);

    public StyleCatalog(Styles styles)
    {
        if (styles == null)
        {
            return;
        }

        foreach (var style in styles.Elements<Style>())
        {
            var id = style.StyleId?.Value;
            if (id != null)
            {
                _stylesById[id] = style;
            }
        }
    }

    public Style GetStyle(string styleId) =>
        styleId != null && _stylesById.TryGetValue(styleId, out var style) ? style : null;

    public string GetStyleName(string styleId) => GetStyle(styleId)?.StyleName?.Val?.Value;
}
