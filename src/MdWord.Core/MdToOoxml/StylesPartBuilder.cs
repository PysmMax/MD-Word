using DocumentFormat.OpenXml.Wordprocessing;

namespace MdWord.Core.MdToOoxml;

/// <summary>
/// Builds the styles.xml shipped with every generated document. Most styles
/// (Normal, HeadingN, Hyperlink, Quote) are style IDs and invariant names
/// only, deliberately with no visual formatting, so that the *target*
/// document's own theme/style definitions (including localized display names
/// such as "Заголовок 1") determine the final appearance when this content
/// is inserted elsewhere. The two code styles are the deliberate exception:
/// Word has no built-in "code" look to inherit, so CodeBlock/CodeInline bake
/// in the monospace font + light shading themselves (per PLAN.md/the 1b
/// brief — formatting only where functionally required).
/// </summary>
internal static class StylesPartBuilder
{
    /// <summary>Fill color for the light-gray code shading (paragraph and run level).</summary>
    public const string CodeShadingFill = "F0F0F0";

    /// <summary>Monospace font used for code blocks/spans.</summary>
    public const string CodeFontName = "Consolas";

    /// <summary>10pt in half-points, OOXML's <c>w:sz</c> unit.</summary>
    public const string CodeFontSizeHalfPoints = "20";

    public static Styles BuildMinimalStyles()
    {
        var styles = new Styles();

        styles.Append(CreateStyle(StyleValues.Paragraph, "Normal", "Normal", isDefault: true));

        for (var level = 1; level <= 6; level++)
        {
            styles.Append(CreateStyle(StyleValues.Paragraph, $"Heading{level}", $"heading {level}"));
        }

        styles.Append(CreateStyle(StyleValues.Character, "Hyperlink", "Hyperlink"));
        styles.Append(CreateStyle(StyleValues.Paragraph, "Quote", "Quote"));
        styles.Append(CreateCodeBlockStyle());
        styles.Append(CreateCodeInlineStyle());

        return styles;
    }

    private static Style CreateStyle(StyleValues type, string styleId, string name, bool isDefault = false)
    {
        var style = new Style
        {
            Type = type,
            StyleId = styleId,
            StyleName = new StyleName { Val = name },
        };

        if (isDefault)
        {
            style.Default = true;
        }

        return style;
    }

    private static Style CreateCodeBlockStyle()
    {
        var style = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "CodeBlock",
            StyleName = new StyleName { Val = "CodeBlock" },
        };

        style.Append(new StyleParagraphProperties(
            new Shading { Val = ShadingPatternValues.Clear, Fill = CodeShadingFill },
            new SpacingBetweenLines { Before = "0", After = "0" }));

        style.Append(new StyleRunProperties(BuildCodeRunFonts(), BuildCodeFontSize()));

        return style;
    }

    private static Style CreateCodeInlineStyle()
    {
        var style = new Style
        {
            Type = StyleValues.Character,
            StyleId = "CodeInline",
            StyleName = new StyleName { Val = "CodeInline" },
        };

        style.Append(new StyleRunProperties(
            BuildCodeRunFonts(),
            BuildCodeFontSize(),
            new Shading { Val = ShadingPatternValues.Clear, Fill = CodeShadingFill }));

        return style;
    }

    private static RunFonts BuildCodeRunFonts() =>
        new() { Ascii = CodeFontName, HighAnsi = CodeFontName, ComplexScript = CodeFontName };

    private static FontSize BuildCodeFontSize() => new() { Val = CodeFontSizeHalfPoints };
}
