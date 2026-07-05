using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace MdWord.Core.Tests;

public class BlockMappingTests
{
    // Mirrors BulletNumId/OrderedNumId (internal to
    // MdWord.Core, not visible here) -- these are the two well-known numId
    // values the numbering.xml part always exposes.
    private const int BulletNumId = 1;
    private const int OrderedNumId = 2;

    private static Body GetBody(string markdown)
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(markdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart.Document.Body;
    }

    [Theory]
    [InlineData("# H1", "Heading1")]
    [InlineData("## H2", "Heading2")]
    [InlineData("### H3", "Heading3")]
    [InlineData("#### H4", "Heading4")]
    [InlineData("##### H5", "Heading5")]
    [InlineData("###### H6", "Heading6")]
    public void Heading_MapsToParagraphWithHeadingStyle(string markdown, string expectedStyleId)
    {
        var body = GetBody(markdown);

        var paragraph = body.Elements<Paragraph>().Single();
        Assert.Equal(expectedStyleId, paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    [Fact]
    public void Heading_UsesSharedInlineMapping_ForEmphasisInHeadingText()
    {
        var body = GetBody("# Bold **word** here");

        var paragraph = body.Elements<Paragraph>().Single();
        var boldRun = paragraph.Elements<Run>().Single(r => r.InnerText == "word");
        Assert.NotNull(boldRun.RunProperties?.Bold);
    }

    [Fact]
    public void ThematicBreak_MapsToParagraphWithBottomBorder()
    {
        var body = GetBody("before\n\n---\n\nafter");

        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(3, paragraphs.Count);

        var breakParagraph = paragraphs[1];
        var border = breakParagraph.ParagraphProperties?.ParagraphBorders?.BottomBorder;
        Assert.NotNull(border);
        Assert.Equal(BorderValues.Single, border.Val.Value);
    }

    [Fact]
    public void QuoteBlock_MapsToParagraphWithQuoteStyle()
    {
        var body = GetBody("> quoted text");

        var paragraph = body.Elements<Paragraph>().Single();
        Assert.Equal("Quote", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal("quoted text", paragraph.InnerText);
    }

    [Fact]
    public void QuoteBlock_WithMultipleParagraphs_MapsEachToItsOwnQuoteStyledParagraph()
    {
        var body = GetBody("> first\n>\n> second");

        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.All(paragraphs, p => Assert.Equal("Quote", p.ParagraphProperties?.ParagraphStyleId?.Val?.Value));
        Assert.Equal("first", paragraphs[0].InnerText);
        Assert.Equal("second", paragraphs[1].InnerText);
    }

    [Fact]
    public void FencedCodeBlock_MapsOneParagraphPerLine_StyledCodeBlock()
    {
        var body = GetBody("```\nline one\nline two\n```");

        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.All(paragraphs, p => Assert.Equal("CodeBlock", p.ParagraphProperties?.ParagraphStyleId?.Val?.Value));
        Assert.Equal("line one", paragraphs[0].InnerText);
        Assert.Equal("line two", paragraphs[1].InnerText);
    }

    [Fact]
    public void FencedCodeBlock_DoesNotApplyMarkdownFormatting_ToItsContent()
    {
        var body = GetBody("```\n*not italic*\n```");

        var paragraph = body.Elements<Paragraph>().Single();
        var run = paragraph.Elements<Run>().Single();
        Assert.Null(run.RunProperties?.Italic);
        Assert.Equal("*not italic*", run.InnerText);
    }

    [Fact]
    public void IndentedCodeBlock_MapsToParagraphsStyledCodeBlock()
    {
        var body = GetBody("before\n\n    indented code\n\nafter");

        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(3, paragraphs.Count);
        Assert.Equal("CodeBlock", paragraphs[1].ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Equal("indented code", paragraphs[1].InnerText);
    }

    [Fact]
    public void BulletList_MapsEachItemToParagraphWithBulletNumId_AtLevelZero()
    {
        var body = GetBody("- one\n- two");

        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(2, paragraphs.Count);
        foreach (var paragraph in paragraphs)
        {
            var numPr = paragraph.ParagraphProperties?.NumberingProperties;
            Assert.NotNull(numPr);
            Assert.Equal(BulletNumId, numPr.NumberingId.Val.Value);
            Assert.Equal(0, numPr.NumberingLevelReference.Val.Value);
        }

        Assert.Equal("one", paragraphs[0].InnerText);
        Assert.Equal("two", paragraphs[1].InnerText);
    }

    [Fact]
    public void OrderedList_MapsEachItemToParagraphWithOrderedNumId_AtLevelZero()
    {
        var body = GetBody("1. one\n2. two");

        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(2, paragraphs.Count);
        foreach (var paragraph in paragraphs)
        {
            var numPr = paragraph.ParagraphProperties?.NumberingProperties;
            Assert.NotNull(numPr);
            Assert.Equal(OrderedNumId, numPr.NumberingId.Val.Value);
            Assert.Equal(0, numPr.NumberingLevelReference.Val.Value);
        }
    }

    [Fact]
    public void NestedList_MapsInnerItems_AtIncrementedLevel_WithOwnListTypeNumId()
    {
        var body = GetBody("- outer\n  1. inner\n- outer two");

        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(3, paragraphs.Count);

        var outer1 = paragraphs[0].ParagraphProperties!.NumberingProperties!;
        Assert.Equal(BulletNumId, outer1.NumberingId!.Val!.Value);
        Assert.Equal(0, outer1.NumberingLevelReference!.Val!.Value);

        var inner = paragraphs[1].ParagraphProperties!.NumberingProperties!;
        Assert.Equal(OrderedNumId, inner.NumberingId!.Val!.Value);
        Assert.Equal(1, inner.NumberingLevelReference!.Val!.Value);

        var outer2 = paragraphs[2].ParagraphProperties!.NumberingProperties!;
        Assert.Equal(BulletNumId, outer2.NumberingId!.Val!.Value);
        Assert.Equal(0, outer2.NumberingLevelReference!.Val!.Value);
    }

    [Fact]
    public void ToOoxml_AddsNumberingDefinitionsPart_WithBulletAndOrderedNumIds()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("- item");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;
        Assert.NotNull(numberingPart);

        var numIds = numberingPart.Numbering.Elements<NumberingInstance>()
            .Select(n => n.NumberID!.Value)
            .ToList();
        Assert.Contains(BulletNumId, numIds);
        Assert.Contains(OrderedNumId, numIds);
    }

    [Fact]
    public void TwoSeparateOrderedLists_ShareOneNumId_KnownPhase1NumberingLimitation()
    {
        // Documents the accepted Phase 1 limitation (see NumberingPartBuilder):
        // two independent top-level ordered lists both get numId=OrderedNumId,
        // so Word will render the second list continuing the first's count
        // (4,5,6 instead of restarting at 1) rather than each starting fresh.
        var body = GetBody("1. a\n2. b\n\ntext between\n\n1. c\n2. d");

        var numIds = body.Elements<Paragraph>()
            .Select(p => p.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value)
            .Where(v => v.HasValue)
            .ToList();

        Assert.Equal(4, numIds.Count);
        Assert.All(numIds, id => Assert.Equal(OrderedNumId, id!.Value));
    }

    [Fact]
    public void ToOoxml_NestedMixedLists_PassOpenXmlValidatorWithZeroErrors()
    {
        var converter = new MarkdownConverter(null);
        var markdown = "- outer bullet\n  1. inner ordered\n     - inner inner bullet\n- outer bullet two\n\n1. top ordered\n2. top ordered two";
        var result = converter.ToOoxml(markdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }

    [Fact]
    public void PipeTable_MapsToWordTable_WithHeaderAndDataRows()
    {
        var body = GetBody("| A | B |\n| --- | --- |\n| one | two |");

        var table = body.Elements<Table>().Single();
        var rows = table.Elements<TableRow>().ToList();
        Assert.Equal(2, rows.Count);

        var headerCells = rows[0].Elements<TableCell>().ToList();
        Assert.Equal(2, headerCells.Count);
        Assert.Equal("A", headerCells[0].InnerText);
        Assert.Equal("B", headerCells[1].InnerText);

        var dataCells = rows[1].Elements<TableCell>().ToList();
        Assert.Equal("one", dataCells[0].InnerText);
        Assert.Equal("two", dataCells[1].InnerText);
    }

    [Fact]
    public void PipeTable_HasExplicitSingleBorders()
    {
        var body = GetBody("| A |\n| --- |\n| one |");

        var table = body.Elements<Table>().Single();
        var borders = table.GetFirstChild<TableProperties>()?.TableBorders;
        Assert.NotNull(borders);
        Assert.Equal(BorderValues.Single, borders.TopBorder.Val.Value);
        Assert.Equal(BorderValues.Single, borders.BottomBorder.Val.Value);
        Assert.Equal(BorderValues.Single, borders.LeftBorder.Val.Value);
        Assert.Equal(BorderValues.Single, borders.RightBorder.Val.Value);
        Assert.Equal(BorderValues.Single, borders.InsideHorizontalBorder.Val.Value);
        Assert.Equal(BorderValues.Single, borders.InsideVerticalBorder.Val.Value);
    }

    [Fact]
    public void PipeTable_HeaderRow_HasBoldRuns()
    {
        var body = GetBody("| A |\n| --- |\n| one |");

        var table = body.Elements<Table>().Single();
        var headerRun = table.Elements<TableRow>().First()
            .Descendants<Run>().Single();
        Assert.NotNull(headerRun.RunProperties?.Bold);

        var dataRun = table.Elements<TableRow>().Last()
            .Descendants<Run>().Single();
        Assert.Null(dataRun.RunProperties?.Bold);
    }

    [Fact]
    public void CodeSpanInHeaderCell_ProducesValidOoxml_WithCorrectRunPropertyOrder()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("| `x` |\n| --- |\n| y |");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }

    [Fact]
    public void ToOoxml_DocumentWithTable_PassesOpenXmlValidatorWithZeroErrors()
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml("| A | B |\n| --- | --- |\n| one | two |\n| **bold** | three |");

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }

    [Fact]
    public void BlockMath_DegradesToParagraphWithDollarDollarDelimiters()
    {
        var body = GetBody("$$\nE=mc^2\n$$");

        var paragraph = body.Elements<Paragraph>().Single();
        Assert.Equal("$$E=mc^2$$", paragraph.InnerText);
    }

    [Fact]
    public void HtmlBlock_DegradesToLiteralTextParagraph_WithoutBeingDropped()
    {
        var body = GetBody("before\n\n<div>raw html</div>\n\nafter");

        var paragraphs = body.Elements<Paragraph>().ToList();
        Assert.Equal(3, paragraphs.Count);
        Assert.Contains("<div>raw html</div>", paragraphs[1].InnerText);
    }
}
