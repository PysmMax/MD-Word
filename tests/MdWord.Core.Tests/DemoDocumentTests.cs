using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace MdWord.Core.Tests;

/// <summary>
/// End-of-Phase-1 sweep: samples/demo.md (PLAN.md §5 — every MVP element in
/// one document) must convert to a valid docx. This is the automated half of
/// the Phase 1 checkpoint; the manual half (open demo.docx in Word, eyeball
/// heading styles/table borders) is the user's, not ours to self-certify.
/// </summary>
public class DemoDocumentTests
{
    private static string DemoMarkdown =>
        File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "samples", "demo.md"));

    [Fact]
    public void Demo_ConvertsAndPassesOpenXmlValidator_WithZeroErrors()
    {
        var converter = new MarkdownConverter(null);

        var result = converter.ToOoxml(DemoMarkdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }

    [Fact]
    public void Demo_ContainsAllSixHeadingLevels()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(DemoMarkdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        var styleIds = body.Elements<Paragraph>()
            .Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
            .Where(v => v != null)
            .ToHashSet();

        for (var level = 1; level <= 6; level++)
        {
            Assert.Contains($"Heading{level}", styleIds);
        }
    }

    [Fact]
    public void Demo_ContainsATableWithThreeColumns()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(DemoMarkdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        var table = body.Elements<Table>().Single();
        var firstRowCellCount = table.Elements<TableRow>().First().Elements<TableCell>().Count();
        Assert.Equal(3, firstRowCellCount);
    }

    [Fact]
    public void Demo_PreservesCyrillicAndEmojiText()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(DemoMarkdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var bodyText = doc.MainDocumentPart.Document.Body.InnerText;

        Assert.Contains("Привіт, світ!", bodyText);
        Assert.Contains("🚀", bodyText);
    }

    [Fact]
    public void Demo_PreservesMarkdownSpecialCharactersAsLiteralText()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(DemoMarkdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var bodyText = doc.MainDocumentPart.Document.Body.InnerText;

        Assert.Contains("* _ | #", bodyText);
    }

    [Fact]
    public void Demo_SoftLineBreakAcrossSourceLines_DoesNotGlueWordsTogether()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(DemoMarkdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var bodyText = doc.MainDocumentPart.Document.Body.InnerText;

        Assert.Contains("текстом, а також", bodyText);
        Assert.DoesNotContain("текстом,а також", bodyText);
    }

    [Fact]
    public void Demo_DegradesFormulasToLiteralDollarText_WithoutLosingContent()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(DemoMarkdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var bodyText = doc.MainDocumentPart.Document.Body.InnerText;

        Assert.Contains("$E=mc^2$", bodyText);
        Assert.Contains("$$", bodyText);
    }
}
