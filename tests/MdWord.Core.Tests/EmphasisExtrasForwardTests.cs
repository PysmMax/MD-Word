using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace MdWord.Core.Tests;

/// <summary>
/// Forward mapping (Markdown → OOXML) for the EmphasisExtras enabled in
/// v1.0.2: subscript (~x~), superscript (^x^), marked (==x==) and inserted
/// (++x++). Mirrors <see cref="InlineMappingTests"/>' fixture style.
/// Delimiter facts verified empirically against Markdig 1.3.2: '~' count 1
/// is subscript, '~' count 2 is strikethrough, '^' is superscript,
/// '=' count 2 is marked, '+' count 2 is inserted; unpaired '~'/'^' in
/// prose produce no emphasis.
/// </summary>
public class EmphasisExtrasForwardTests
{
    private static Paragraph SingleParagraph(string markdown)
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(markdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        return body.Elements<Paragraph>().Single();
    }

    [Fact]
    public void Subscript_SingleTilde_MapsToVertAlignSubscript()
    {
        var paragraph = SingleParagraph("H~2~O");

        var run = paragraph.Elements<Run>().Single(r => r.InnerText == "2");
        var vertAlign = run.RunProperties?.VerticalTextAlignment;
        Assert.NotNull(vertAlign);
        Assert.Equal(VerticalPositionValues.Subscript, vertAlign.Val.Value);
    }

    [Fact]
    public void Superscript_SingleCaret_MapsToVertAlignSuperscript()
    {
        var paragraph = SingleParagraph("E = mc^2^");

        var run = paragraph.Elements<Run>().Single(r => r.InnerText == "2");
        var vertAlign = run.RunProperties?.VerticalTextAlignment;
        Assert.NotNull(vertAlign);
        Assert.Equal(VerticalPositionValues.Superscript, vertAlign.Val.Value);
    }

    [Fact]
    public void Strikethrough_DoubleTilde_StillMapsToStrike_NotSubscript()
    {
        var paragraph = SingleParagraph("~~gone~~");

        var run = paragraph.Elements<Run>().Single();
        Assert.NotNull(run.RunProperties?.Strike);
        Assert.Null(run.RunProperties?.VerticalTextAlignment);
    }

    [Fact]
    public void Marked_DoubleEquals_MapsToYellowHighlight()
    {
        var paragraph = SingleParagraph("==note==");

        var run = paragraph.Elements<Run>().Single();
        var highlight = run.RunProperties?.Highlight;
        Assert.NotNull(highlight);
        Assert.Equal(HighlightColorValues.Yellow, highlight.Val.Value);
    }

    [Fact]
    public void Inserted_DoublePlus_MapsToSingleUnderline()
    {
        var paragraph = SingleParagraph("++added++");

        var run = paragraph.Elements<Run>().Single();
        var underline = run.RunProperties?.Underline;
        Assert.NotNull(underline);
        Assert.Equal(UnderlineValues.Single, underline.Val.Value);
    }

    [Fact]
    public void UnpairedTildeAndCaret_StayLiteralText()
    {
        var paragraph = SingleParagraph("range 5~10 or 3^4 alone");

        Assert.Equal("range 5~10 or 3^4 alone", paragraph.InnerText);
    }

    [Fact]
    public void BoldSubscript_CarriesBothProperties()
    {
        var paragraph = SingleParagraph("**H~2~O**");

        var run = paragraph.Elements<Run>().Single(r => r.InnerText == "2");
        Assert.NotNull(run.RunProperties?.Bold);
        var vertAlign = run.RunProperties?.VerticalTextAlignment;
        Assert.NotNull(vertAlign);
        Assert.Equal(VerticalPositionValues.Subscript, vertAlign.Val.Value);
    }

    [Fact]
    public void CombinedExtras_ProduceValidOoxml()
    {
        // Also guards the CT_RPr child order (b, i, strike, highlight, u,
        // vertAlign) — OpenXmlValidator flags out-of-sequence rPr children.
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("**bold ==mark== with H~2~O and x^2^ and ++ins++**");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }
}
