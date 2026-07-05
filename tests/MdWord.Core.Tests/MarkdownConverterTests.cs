using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace MdWord.Core.Tests;

public class MarkdownConverterTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // ---- Increment 1: core plumbing (empty-ish document) ----

    [Fact]
    public void ToOoxml_FlatOpcRoundTrips_ViaFromFlatOpcString()
    {
        var converter = new MarkdownConverter(null);

        var result = converter.ToOoxml(string.Empty);

        using var reopened = WordprocessingDocument.FromFlatOpcString(result.FlatOpc);
        Assert.NotNull(reopened.MainDocumentPart);
        Assert.NotNull(reopened.MainDocumentPart.Document.Body);
    }

    [Fact]
    public void ToOoxml_ReturnsBothDocxBytesAndFlatOpc_ForSameContent()
    {
        var converter = new MarkdownConverter(null);

        var result = converter.ToOoxml("Hello");

        Assert.NotNull(result.DocxBytes);
        Assert.True(result.DocxBytes.Length > 0);
        Assert.False(string.IsNullOrEmpty(result.FlatOpc));

        using var fromBytes = WordprocessingDocument.Open(new MemoryStream(result.DocxBytes), false);
        using var fromFlatOpc = WordprocessingDocument.FromFlatOpcString(result.FlatOpc);

        var textFromBytes = fromBytes.MainDocumentPart.Document.Body.InnerText;
        var textFromFlatOpc = fromFlatOpc.MainDocumentPart.Document.Body.InnerText;
        Assert.Equal(textFromBytes, textFromFlatOpc);
        Assert.Equal("Hello", textFromBytes);
    }

    [Fact]
    public void ToOoxml_GeneratesMinimalStylesPart_WithExpectedStyleIds()
    {
        var converter = new MarkdownConverter(null);

        var result = converter.ToOoxml(string.Empty);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;
        Assert.NotNull(stylesPart);

        var stylesXml = stylesPart.Styles.OuterXml;
        var stylesXDoc = XDocument.Parse(stylesXml);

        var styleIds = stylesXDoc.Root!.Elements(W + "style")
            .Select(s => (string)s.Attribute(W + "styleId"))
            .ToArray();

        Assert.Contains("Normal", styleIds);
        Assert.Contains("Heading1", styleIds);
        Assert.Contains("Heading2", styleIds);
        Assert.Contains("Heading3", styleIds);
        Assert.Contains("Heading4", styleIds);
        Assert.Contains("Heading5", styleIds);
        Assert.Contains("Heading6", styleIds);
        Assert.Contains("Hyperlink", styleIds);
        Assert.Contains("Quote", styleIds);
        Assert.Contains("CodeBlock", styleIds);
        Assert.Contains("CodeInline", styleIds);

        // Invariant w:name for headings, e.g. styleId="Heading1" -> name val="heading 1"
        foreach (var level in new[] { 1, 2, 3, 4, 5, 6 })
        {
            var style = stylesXDoc.Root!.Elements(W + "style")
                .First(s => (string)s.Attribute(W + "styleId") == $"Heading{level}");
            var nameVal = (string)style.Element(W + "name")!.Attribute(W + "val");
            Assert.Equal($"heading {level}", nameVal);
        }

        // Hyperlink must be a *character* style, everything else a paragraph style
        // (1c's link mapping depends on this).
        string StyleType(string styleId) =>
            (string)stylesXDoc.Root!.Elements(W + "style")
                .First(s => (string)s.Attribute(W + "styleId") == styleId)
                .Attribute(W + "type");

        Assert.Equal("character", StyleType("Hyperlink"));
        Assert.Equal("paragraph", StyleType("Normal"));
        Assert.Equal("paragraph", StyleType("Quote"));
        Assert.Equal("paragraph", StyleType("Heading1"));
        Assert.Equal("paragraph", StyleType("CodeBlock"));
        Assert.Equal("character", StyleType("CodeInline"));
    }

    [Fact]
    public void ToOoxml_EmptyMarkdown_ProducesMinimalDocument_PassingOpenXmlValidator()
    {
        var converter = new MarkdownConverter(null);

        var result = converter.ToOoxml(string.Empty);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }

    [Fact]
    public void ToOoxml_GeneratedBody_HasNoSectionProperties()
    {
        var converter = new MarkdownConverter(null);

        var result = converter.ToOoxml("Some text");

        var bodyXml = XDocument.Parse(result.FlatOpc);
        var sectPr = bodyXml.Descendants(W + "body")
            .Elements(W + "sectPr");
        Assert.Empty(sectPr);
    }

    [Fact]
    public void ToOoxml_GeneratedDocument_PassesOpenXmlValidatorWithZeroErrors()
    {
        var converter = new MarkdownConverter(null);

        var result = converter.ToOoxml("A simple paragraph.");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }

    // ---- Increment 2: single plain-text paragraph content mapping ----

    [Fact]
    public void ToOoxml_PlainTextParagraph_MapsToSingleNormalStyledParagraph()
    {
        var converter = new MarkdownConverter(null);

        var result = converter.ToOoxml("Hello world");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        var paragraphs = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToList();
        Assert.Single(paragraphs);

        var pStyle = paragraphs[0].ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        Assert.Equal("Normal", pStyle);
        Assert.Equal("Hello world", paragraphs[0].InnerText);
    }

    [Fact]
    public void ToOoxml_CyrillicParagraph_RoundTripsTextExactly()
    {
        const string cyrillic = "Привіт, світ! Це напис українською.";
        var converter = new MarkdownConverter(null);

        var result = converter.ToOoxml(cyrillic);

        using var fromBytes = WordprocessingDocument.Open(new MemoryStream(result.DocxBytes), false);
        using var fromFlatOpc = WordprocessingDocument.FromFlatOpcString(result.FlatOpc);

        Assert.Equal(cyrillic, fromBytes.MainDocumentPart.Document.Body.InnerText);
        Assert.Equal(cyrillic, fromFlatOpc.MainDocumentPart.Document.Body.InnerText);
    }

    // ---- ToMarkdown (Phase 4): null-guard mirrors ToOoxml's own contract ----

    [Fact]
    public void ToMarkdown_Null_ThrowsConvertException()
    {
        var converter = new MarkdownConverter(null);

        Assert.Throws<ConvertException>(() => converter.ToMarkdown(null));
    }

    [Fact]
    public void ToMarkdown_MalformedFlatOpc_ThrowsConvertException()
    {
        var converter = new MarkdownConverter(null);

        Assert.Throws<ConvertException>(() => converter.ToMarkdown("not a valid flat opc document"));
    }

    [Fact]
    public void ToOoxml_ThenToMarkdown_PlainParagraph_RoundTrips()
    {
        var converter = new MarkdownConverter(null);

        var forward = converter.ToOoxml("Hello world");
        var reverse = converter.ToMarkdown(forward.FlatOpc);

        Assert.Contains("Hello world", reverse.Markdown);
    }
}
