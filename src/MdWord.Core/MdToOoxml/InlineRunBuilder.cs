using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Extensions.Mathematics;
using Markdig.Syntax.Inlines;
using MdWord.Core.Math;
using MathOfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;

namespace MdWord.Core.MdToOoxml;

/// <summary>
/// Maps a Markdig inline tree to a flat sequence of OOXML run-level elements
/// (<see cref="Run"/>, later <see cref="Hyperlink"/>). Shared by every block
/// mapper (paragraphs, headings, quotes, list items, table cells) so that
/// emphasis/code/links behave identically everywhere they appear.
///
/// Formatting (bold/italic/strike) is tracked as accumulated flags while
/// walking nested <see cref="EmphasisInline"/> nodes, rather than switched on
/// a single node's delimiter count. This is deliberate: Markdig may represent
/// "***x***" as one EmphasisInline (DelimiterCount 3) or as two nested ones
/// depending on the surrounding delimiter run, and accumulating flags through
/// recursion handles both shapes without caring which one occurred.
/// </summary>
internal static class InlineRunBuilder
{
    private readonly struct Format
    {
        public readonly bool Bold;
        public readonly bool Italic;
        public readonly bool Strike;
        public readonly bool Subscript;
        public readonly bool Superscript;
        public readonly bool Highlight;
        public readonly bool Underline;

        private Format(bool bold, bool italic, bool strike, bool subscript, bool superscript, bool highlight, bool underline)
        {
            Bold = bold;
            Italic = italic;
            Strike = strike;
            Subscript = subscript;
            Superscript = superscript;
            Highlight = highlight;
            Underline = underline;
        }

        public Format WithBold() => new(true, Italic, Strike, Subscript, Superscript, Highlight, Underline);
        public Format WithItalic() => new(Bold, true, Strike, Subscript, Superscript, Highlight, Underline);
        public Format WithStrike() => new(Bold, Italic, true, Subscript, Superscript, Highlight, Underline);
        public Format WithSubscript() => new(Bold, Italic, Strike, true, Superscript, Highlight, Underline);
        public Format WithSuperscript() => new(Bold, Italic, Strike, Subscript, true, Highlight, Underline);
        public Format WithHighlight() => new(Bold, Italic, Strike, Subscript, Superscript, true, Underline);
        public Format WithUnderline() => new(Bold, Italic, Strike, Subscript, Superscript, Highlight, true);
    }

    public static IEnumerable<OpenXmlElement> Build(ContainerInline container, MainDocumentPart mainPart = null, MathConversionContext mathContext = null)
    {
        var result = new List<OpenXmlElement>();
        BuildInto(container, default, result, mainPart, mathContext);
        return result;
    }

    private static void BuildInto(ContainerInline container, Format format, List<OpenXmlElement> result, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        foreach (var inline in container)
        {
            switch (inline)
            {
                case LiteralInline literal:
                    result.Add(BuildRun(literal.Content.ToString(), format));
                    break;
                case LineBreakInline { IsHard: true }:
                    result.Add(BuildBreakRun(format));
                    break;
                case LineBreakInline:
                    // Soft line break: CommonMark renders this as whitespace, not
                    // a line break. Dropping it entirely (matching no case below)
                    // would silently glue the words on either side together.
                    result.Add(BuildRun(" ", format));
                    break;
                case HtmlInline html:
                    // A literal <br> is the convention our own copy path emits for
                    // a line break inside a table cell — turn it back into a real
                    // Word break. All other raw HTML still degrades to its literal
                    // source text (the initial plan, Phase 1: "HtmlInline/HtmlBlock → literal").
                    if (IsLineBreakTag(html.Tag))
                    {
                        result.Add(BuildBreakRun(format));
                    }
                    else
                    {
                        result.Add(BuildRun(html.Tag, format));
                    }
                    break;
                case HtmlEntityInline entity:
                    result.Add(BuildRun(entity.Transcoded.ToString(), format));
                    break;
                case CodeInline code:
                    result.Add(BuildCodeRun(code.Content));
                    break;
                case MathInline math:
                    result.Add(BuildMathInlineElement(math, format, mathContext));
                    break;
                case LinkInline { IsImage: true } image:
                    // Image embedding is not supported -- degrade to the alt text
                    // plus a warning so content is never silently lost.
                    result.Add(BuildRun(GetPlainText(image), format));
                    mathContext?.Warnings.Add($"Image '{image.Url}' skipped — inserted the caption only.");
                    break;
                case LinkInline link:
                    result.Add(BuildHyperlink(link, format, mainPart, mathContext));
                    break;
                case AutolinkInline autolink:
                    result.Add(BuildAutolinkHyperlink(autolink, format, mainPart, mathContext));
                    break;
                case EmphasisInline emphasis:
                    BuildInto(emphasis, ApplyEmphasis(format, emphasis), result, mainPart, mathContext);
                    break;
                case ContainerInline nested:
                    BuildInto(nested, format, result, mainPart, mathContext);
                    break;
            }
        }
    }

    private static bool IsLineBreakTag(string tag)
    {
        if (string.IsNullOrEmpty(tag))
        {
            return false;
        }

        var normalized = tag.Replace(" ", string.Empty);
        return normalized.Equals("<br>", System.StringComparison.OrdinalIgnoreCase)
            || normalized.Equals("<br/>", System.StringComparison.OrdinalIgnoreCase);
    }

    private static OpenXmlElement BuildMathInlineElement(MathInline math, Format format, MathConversionContext mathContext)
    {
        var tex = math.Content.ToString();

        if (mathContext != null && mathContext.IsActive)
        {
            if (MathConverter.TryConvert(tex, displayMode: false, mathContext.XslPaths, out var ommlOuterXml, out var failureReason))
            {
                return new MathOfficeMath(ommlOuterXml);
            }

            mathContext.Warnings.Add($"Formula `{tex}` inserted as text: {failureReason}");
        }

        // No xslPaths (or conversion failed for this one formula): degrade to
        // the literal "$...$" source so content is never silently lost.
        return BuildRun($"${tex}$", format);
    }

    private static Format ApplyEmphasis(Format format, EmphasisInline emphasis)
    {
        switch (emphasis.DelimiterChar)
        {
            case '~':
                // Markdig EmphasisExtras share the tilde: '~~' (count 2) is
                // strikethrough, a single '~' (count 1) is subscript.
                return emphasis.DelimiterCount >= 2 ? format.WithStrike() : format.WithSubscript();
            case '^':
                return format.WithSuperscript();
            case '=':
                return format.WithHighlight();
            case '+':
                return format.WithUnderline();
        }

        return emphasis.DelimiterCount switch
        {
            2 => format.WithBold(),
            3 => format.WithBold().WithItalic(),
            _ => format.WithItalic(),
        };
    }

    private static Run BuildRun(string text, Format format)
    {
        var run = new Run();

        var runProperties = BuildRunProperties(format);
        if (runProperties.HasChildren)
        {
            run.RunProperties = runProperties;
        }

        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        return run;
    }

    private static Run BuildBreakRun(Format format)
    {
        var run = new Run();

        var runProperties = BuildRunProperties(format);
        if (runProperties.HasChildren)
        {
            run.RunProperties = runProperties;
        }

        run.Append(new Break());

        return run;
    }

    private static Run BuildCodeRun(string text)
    {
        var run = new Run(new RunProperties(new RunStyle { Val = "CodeInline" }));
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static Hyperlink BuildHyperlink(LinkInline link, Format format, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        var runs = new List<OpenXmlElement>();
        BuildInto(link, format, runs, mainPart, mathContext);

        foreach (var element in runs)
        {
            if (element is Run run)
            {
                run.RunProperties ??= new RunProperties();

                // A code span already carries rStyle=CodeInline (maxOccurs=1 in the
                // schema); don't clobber it with a second rStyle for the link. The
                // run stays code-styled but is still clickable via the w:hyperlink
                // wrapper -- a reasonable degradation.
                run.RunProperties.RunStyle ??= new RunStyle { Val = "Hyperlink" };
            }
        }

        var hyperlink = new Hyperlink(runs);

        if (mainPart != null)
        {
            // A malformed destination (Markdig's <...> form allows spaces) must
            // degrade this one link, not abort the whole document (the initial plan §6).
            if (System.Uri.TryCreate(link.Url ?? string.Empty, System.UriKind.RelativeOrAbsolute, out var uri))
            {
                if (uri.IsAbsoluteUri && !IsAllowedScheme(uri.Scheme))
                {
                    // SEC-02: only http(s)/mailto destinations get a live hyperlink --
                    // anything else (file://, javascript:, custom schemes, ...) degrades
                    // to plain text plus a warning rather than a silently created link.
                    mathContext?.Warnings.Add($"A link with scheme '{uri.Scheme}' skipped — inserted the text only ('{link.Url}').");
                }
                else
                {
                    var relationship = mainPart.AddHyperlinkRelationship(uri, true);
                    hyperlink.Id = relationship.Id;
                }
            }
            else
            {
                mathContext?.Warnings.Add($"A link with an invalid address '{link.Url}' inserted without a hyperlink.");
            }
        }

        return hyperlink;
    }

    private static Hyperlink BuildAutolinkHyperlink(AutolinkInline autolink, Format format, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        var run = BuildRun(autolink.Url, format);
        run.RunProperties ??= new RunProperties();
        run.RunProperties.PrependChild(new RunStyle { Val = "Hyperlink" });

        var hyperlink = new Hyperlink(run);

        if (mainPart != null)
        {
            var url = autolink.IsEmail ? "mailto:" + autolink.Url : autolink.Url;
            if (System.Uri.TryCreate(url ?? string.Empty, System.UriKind.RelativeOrAbsolute, out var uri))
            {
                if (uri.IsAbsoluteUri && !IsAllowedScheme(uri.Scheme))
                {
                    mathContext?.Warnings.Add($"A link with scheme '{uri.Scheme}' skipped — inserted the text only ('{autolink.Url}').");
                }
                else
                {
                    var relationship = mainPart.AddHyperlinkRelationship(uri, true);
                    hyperlink.Id = relationship.Id;
                }
            }
        }

        return hyperlink;
    }

    private static bool IsAllowedScheme(string scheme) =>
        string.Equals(scheme, System.Uri.UriSchemeHttp, System.StringComparison.OrdinalIgnoreCase)
        || string.Equals(scheme, System.Uri.UriSchemeHttps, System.StringComparison.OrdinalIgnoreCase)
        || string.Equals(scheme, System.Uri.UriSchemeMailto, System.StringComparison.OrdinalIgnoreCase);

    private static string GetPlainText(ContainerInline container)
    {
        var builder = new System.Text.StringBuilder();
        foreach (var inline in container)
        {
            if (inline is LiteralInline literal)
            {
                builder.Append(literal.Content.ToString());
            }
            else if (inline is ContainerInline nested)
            {
                builder.Append(GetPlainText(nested));
            }
        }

        return builder.ToString();
    }

    private static RunProperties BuildRunProperties(Format format)
    {
        // Child order matters: CT_RPr is an xsd:sequence (…, b, i, …,
        // strike, …, highlight, u, …, vertAlign, …) and OpenXmlValidator
        // rejects out-of-order children — keep appends in schema order.
        var runProperties = new RunProperties();
        if (format.Bold)
        {
            runProperties.Append(new Bold());
        }

        if (format.Italic)
        {
            runProperties.Append(new Italic());
        }

        if (format.Strike)
        {
            runProperties.Append(new Strike());
        }

        if (format.Highlight)
        {
            runProperties.Append(new Highlight { Val = HighlightColorValues.Yellow });
        }

        if (format.Underline)
        {
            runProperties.Append(new Underline { Val = UnderlineValues.Single });
        }

        if (format.Superscript)
        {
            runProperties.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        }
        else if (format.Subscript)
        {
            runProperties.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
        }

        return runProperties;
    }
}
