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
    public void Link_WithMailtoScheme_MapsToHyperlinkElement_WithMailtoRelationship()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("[x](mailto:a@b.com)");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var mainPart = doc.MainDocumentPart;
        var body = mainPart.Document.Body;

        var hyperlink = body.Descendants<Hyperlink>().Single();
        var relationshipId = hyperlink.Id;
        Assert.False(string.IsNullOrEmpty(relationshipId));
        var rel = mainPart.HyperlinkRelationships.Single(r => r.Id == relationshipId);
        Assert.Equal("mailto:a@b.com", rel.Uri.ToString());
    }

    // SEC-02 regression (security audit): Uri.TryCreate(..., UriKind.RelativeOrAbsolute, ...)
    // parses a scheme-relative "//host/share" as a RELATIVE uri (IsAbsoluteUri == false),
    // even though Word may resolve it as a UNC network path when clicked. The old check
    // ("uri.IsAbsoluteUri && !IsAllowedScheme(...)") only fired for absolute uris, so this
    // -- along with every other non-absolute destination -- fell straight through to a live
    // hyperlink, bypassing the http/https/mailto allow-list entirely.
    //
    // "\\host\share" is included here too, not in the "already blocked" regression group
    // below: CommonMark processes backslash-escapes in link destinations, and "\\" is
    // itself an escapable character, so this markdown source decodes to link.Url ==
    // @"\host\share" (one leading backslash, not two). A single leading backslash is not
    // .NET's UNC marker (only "\\" is), so Uri.TryCreate parses it as a RELATIVE uri --
    // same hole as "//host/share", not a pre-existing disallowed-absolute-scheme case.
    [Theory]
    [InlineData("//host/share")]
    [InlineData("../../x.exe")]
    [InlineData("x.exe")]
    [InlineData("#section")]
    [InlineData(@"\\host\share")]
    public void Link_WithNonAbsoluteTarget_DoesNotCreateRelationship_ButKeepsText(string url)
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml($"[x]({url})");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        var hyperlink = body.Descendants<Hyperlink>().Single();
        Assert.True(string.IsNullOrEmpty(hyperlink.Id));
        Assert.Contains("x", body.InnerText);
        Assert.Contains(result.Warnings, w => w.Contains("relative or unresolved target"));
    }

    // Regression: already blocked before the fix (parses as an absolute uri with a
    // disallowed "file" scheme) and must stay blocked. Unlike "\\host\share" above, the
    // single backslash before "x.exe" is not a CommonMark escape target (backslash before
    // a non-punctuation character survives literally), so this decodes to exactly
    // @"C:\x.exe" and .NET recognizes the drive-letter form as an absolute file:// uri.
    [Fact]
    public void Link_WithDisallowedAbsoluteTarget_StillDoesNotCreateRelationship()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(@"[x](C:\x.exe)");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        var hyperlink = body.Descendants<Hyperlink>().Single();
        Assert.True(string.IsNullOrEmpty(hyperlink.Id));
        Assert.Contains("x", body.InnerText);
        Assert.Contains(result.Warnings, w => w.Contains("scheme 'file'"));
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

    [Fact]
    public void Autolink_Email_MapsToHyperlinkElement_WithMailtoRelationship()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("see <a@b.com> here");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var mainPart = doc.MainDocumentPart;
        var body = mainPart.Document.Body;

        var hyperlink = body.Descendants<Hyperlink>().Single();
        var relationshipId = hyperlink.Id;
        Assert.False(string.IsNullOrEmpty(relationshipId));
        var rel = mainPart.HyperlinkRelationships.Single(r => r.Id == relationshipId);
        Assert.Equal("mailto:a@b.com", rel.Uri.ToString());
    }

    // SEC-02 regression (security audit), autolink side: same allow-list fix as the
    // explicit-link case above, applied to BuildAutolinkHyperlink. A disallowed absolute
    // scheme was already blocked before the fix; this confirms it still is.
    [Fact]
    public void Autolink_WithDisallowedScheme_DoesNotCreateRelationship_ButKeepsText()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("see <file://host/share> here");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        var hyperlink = body.Descendants<Hyperlink>().Single();
        Assert.True(string.IsNullOrEmpty(hyperlink.Id));
        Assert.Contains("file://host/share", body.InnerText);
        Assert.Contains(result.Warnings, w => w.Contains("file"));
    }
}
