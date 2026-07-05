using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace MdWord.Core.Tests;

public class InlineMappingTests
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
    public void Italic_SingleAsterisk_MapsToRunWithItalicProperty()
    {
        var paragraph = SingleParagraph("*italic*");

        var run = paragraph.Elements<Run>().Single();
        Assert.NotNull(run.RunProperties?.Italic);
        Assert.Equal("italic", run.InnerText);
    }

    [Fact]
    public void Bold_DoubleAsterisk_MapsToRunWithBoldProperty()
    {
        var paragraph = SingleParagraph("**bold**");

        var run = paragraph.Elements<Run>().Single();
        Assert.NotNull(run.RunProperties?.Bold);
        Assert.Null(run.RunProperties?.Italic);
        Assert.Equal("bold", run.InnerText);
    }

    [Fact]
    public void BoldItalic_TripleAsterisk_MapsToRunWithBothProperties()
    {
        var paragraph = SingleParagraph("***both***");

        var run = paragraph.Elements<Run>().Single();
        Assert.NotNull(run.RunProperties?.Bold);
        Assert.NotNull(run.RunProperties?.Italic);
        Assert.Equal("both", run.InnerText);
    }

    [Fact]
    public void Strikethrough_DoubleTilde_MapsToRunWithStrikeProperty()
    {
        var paragraph = SingleParagraph("~~gone~~");

        var run = paragraph.Elements<Run>().Single();
        Assert.NotNull(run.RunProperties?.Strike);
        Assert.Equal("gone", run.InnerText);
    }

    [Fact]
    public void CodeSpan_MapsToRunWithCodeInlineStyle()
    {
        var paragraph = SingleParagraph("plain `code` plain");

        var runs = paragraph.Elements<Run>().ToList();
        var codeRun = runs.Single(r => r.InnerText == "code");
        Assert.Equal("CodeInline", codeRun.RunProperties?.RunStyle?.Val?.Value);
    }

    [Fact]
    public void HardLineBreak_MapsToRunWithBreak()
    {
        var paragraph = SingleParagraph("line one  \nline two");

        var breaks = paragraph.Descendants<Break>().ToList();
        Assert.Single(breaks);
    }

    [Fact]
    public void Link_MapsToHyperlinkElement_WithExternalRelationship()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("[click me](https://example.com/)");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var mainPart = doc.MainDocumentPart;
        var body = mainPart.Document.Body;

        var hyperlink = body.Descendants<Hyperlink>().Single();
        var run = hyperlink.Elements<Run>().Single();
        Assert.Equal("click me", run.InnerText);
        Assert.Equal("Hyperlink", run.RunProperties?.RunStyle?.Val?.Value);

        var relationshipId = hyperlink.Id;
        Assert.False(string.IsNullOrEmpty(relationshipId));
        var rel = mainPart.HyperlinkRelationships.Single(r => r.Id == relationshipId);
        Assert.Equal("https://example.com/", rel.Uri.ToString());
    }

    [Fact]
    public void Link_WithDisallowedScheme_DoesNotCreateRelationship_ButKeepsText()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("[share](file://host/share)");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        var hyperlink = body.Descendants<Hyperlink>().Single();
        Assert.True(string.IsNullOrEmpty(hyperlink.Id));
        Assert.Contains("share", body.InnerText);
        Assert.Contains(result.Warnings, w => w.Contains("file"));
    }

    [Fact]
    public void InlineMath_DegradesToLiteralTextWithDollarDelimiters()
    {
        var paragraph = SingleParagraph("before $E=mc^2$ after");

        var runs = paragraph.Elements<Run>().ToList();
        Assert.Contains(runs, r => r.InnerText == "$E=mc^2$");
    }

    [Fact]
    public void SoftLineBreak_DoesNotConcatenateAdjacentWordsWithoutASpace()
    {
        var paragraph = SingleParagraph("a\nb");

        Assert.Equal("a b", paragraph.InnerText);
    }

    [Fact]
    public void HtmlInline_DegradesToLiteralTagText()
    {
        // <br/> is excluded from this generic case since Task 1.10 (LIVE-1)
        // reclaims it as a real line break -- any other inline tag still
        // degrades to its literal source text.
        var paragraph = SingleParagraph("before <span> after");

        Assert.Contains("<span>", paragraph.InnerText);
    }

    [Fact]
    public void HtmlEntity_DegradesToLiteralTranscodedText()
    {
        var paragraph = SingleParagraph("a &mdash; b");

        Assert.Contains("—", paragraph.InnerText);
    }

    [Fact]
    public void CodeSpanInsideLink_ProducesValidOoxml_WithNoDuplicateRunStyle()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("[`x`](https://example.com/)");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }

    [Fact]
    public void Autolink_MapsToHyperlinkElement_WithUrlAsVisibleText()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("see <https://example.com/> here");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        var hyperlink = body.Descendants<Hyperlink>().Single();
        Assert.Equal("https://example.com/", hyperlink.InnerText);
    }
}
