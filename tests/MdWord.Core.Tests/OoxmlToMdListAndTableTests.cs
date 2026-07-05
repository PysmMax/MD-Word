using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace MdWord.Core.Tests;

/// <summary>
/// Phase 4 reverse mapping: lists (incl. a foreign, arbitrary <c>numId</c>)
/// and tables (incl. <c>gridSpan</c>/<c>vMerge</c> degrade-to-empty-cell).
/// </summary>
public class OoxmlToMdListAndTableTests
{
    private static MarkdownResult ToMarkdown(string markdown)
    {
        var converter = new MarkdownConverter(null);
        var forward = converter.ToOoxml(markdown);
        return converter.ToMarkdown(forward.FlatOpc);
    }

    private static string BuildFlatOpc(Body body, Numbering numbering = null)
    {
        using var stream = new MemoryStream();
        string flatOpc;

        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(body);

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();
            stylesPart.Styles.Save();

            if (numbering != null)
            {
                var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = numbering;
                numberingPart.Numbering.Save();
            }

            mainPart.Document.Save();
            flatOpc = doc.ToFlatOpcString();
        }

        return flatOpc;
    }

    // ---- Lists: own bullet/ordered numIds ----

    [Fact]
    public void BulletList_RoundTrips()
    {
        var reverse = ToMarkdown("- one\n- two\n- three");

        Assert.Contains("- one", reverse.Markdown);
        Assert.Contains("- two", reverse.Markdown);
        Assert.Contains("- three", reverse.Markdown);
    }

    [Fact]
    public void OrderedList_RoundTrips()
    {
        var reverse = ToMarkdown("1. first\n2. second\n3. third");

        Assert.Contains("1. first", reverse.Markdown);
        Assert.Contains("2. second", reverse.Markdown);
        Assert.Contains("3. third", reverse.Markdown);
    }

    [Fact]
    public void NestedList_RoundTrips_WithIndentation()
    {
        var reverse = ToMarkdown("- outer\n  1. inner one\n  2. inner two");

        Assert.Contains("- outer", reverse.Markdown);
        Assert.Contains("1. inner one", reverse.Markdown);
        Assert.Contains("2. inner two", reverse.Markdown);
    }

    // ---- Lists: foreign document, arbitrary numId ----

    [Fact]
    public void ForeignNumId_BulletLevel_ResolvesViaNumberingPart()
    {
        // numId=42 (never used by our own generator, which only ever emits
        // 1/2) -> abstractNumId=7, level 0 = Bullet.
        var numbering = new Numbering(
            new AbstractNum(
                new Level(new NumberingFormat { Val = NumberFormatValues.Bullet }, new LevelText { Val = "*" }) { LevelIndex = 0 })
            { AbstractNumberId = 7 },
            new NumberingInstance(new AbstractNumId { Val = 7 }) { NumberID = 42 });

        var body = new Body(
            new Paragraph(
                new ParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 42 })),
                new Run(new Text("foreign bullet item"))));

        var flatOpc = BuildFlatOpc(body, numbering);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("- foreign bullet item", reverse.Markdown);
    }

    [Fact]
    public void ForeignNumId_OrderedLevel_ResolvesViaNumberingPart()
    {
        var numbering = new Numbering(
            new AbstractNum(
                new Level(new NumberingFormat { Val = NumberFormatValues.Decimal }, new LevelText { Val = "%1." }) { LevelIndex = 0 })
            { AbstractNumberId = 9 },
            new NumberingInstance(new AbstractNumId { Val = 9 }) { NumberID = 77 });

        var body = new Body(
            new Paragraph(
                new ParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 77 })),
                new Run(new Text("foreign ordered item"))));

        var flatOpc = BuildFlatOpc(body, numbering);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("1. foreign ordered item", reverse.Markdown);
    }

    // ---- Tables ----

    [Fact]
    public void SimpleTable_RoundTrips_AsPipeTableWithHeaderSeparator()
    {
        var reverse = ToMarkdown("| A | B |\n| --- | --- |\n| 1 | 2 |");

        // The forward direction bolds header-row cell text (BlockWalker.
        // BuildTableCellParagraph) -- reading it back faithfully reproduces
        // that bold run, hence **A**/**B**, not bare "A"/"B".
        Assert.Contains("| **A** | **B** |", reverse.Markdown);
        Assert.Contains("| --- | --- |", reverse.Markdown);
        Assert.Contains("| 1 | 2 |", reverse.Markdown);
    }

    [Fact]
    public void Table_WithGridSpan_DegradesToEmptyCell_WithWarning()
    {
        var body = new Body(
            new Table(
                new TableGrid(new GridColumn(), new GridColumn(), new GridColumn()),
                new TableRow(
                    new TableCell(new TableCellProperties(new GridSpan { Val = 2 }), new Paragraph(new Run(new Text("spans two")))),
                    new TableCell(new Paragraph(new Run(new Text("normal"))))),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("r2c1")))),
                    new TableCell(new Paragraph(new Run(new Text("r2c2")))),
                    new TableCell(new Paragraph(new Run(new Text("r2c3")))))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("spans two", reverse.Markdown);
        Assert.NotEmpty(reverse.Warnings);
        Assert.Contains(reverse.Warnings, w => w.Contains("gridSpan"));

        // 3 grid columns => every data row line must have exactly 3 cells,
        // even the gridSpan row (spans-two, empty, normal).
        var lines = reverse.Markdown.Split('\n');
        var firstRowLine = System.Array.Find(lines, l => l.Contains("spans two"));
        Assert.NotNull(firstRowLine);
        Assert.Equal(3, firstRowLine.Split('|').Length - 2);
    }

    [Fact]
    public void Table_WithVerticalMerge_DegradesToEmptyCell_WithWarning()
    {
        var body = new Body(
            new Table(
                new TableGrid(new GridColumn(), new GridColumn()),
                new TableRow(
                    new TableCell(new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Restart }), new Paragraph(new Run(new Text("merged")))),
                    new TableCell(new Paragraph(new Run(new Text("normal1"))))),
                new TableRow(
                    new TableCell(new TableCellProperties(new VerticalMerge()), new Paragraph()),
                    new TableCell(new Paragraph(new Run(new Text("normal2")))))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("merged", reverse.Markdown);
        Assert.NotEmpty(reverse.Warnings);
        Assert.Contains(reverse.Warnings, w => w.Contains("vMerge"));
    }
}
