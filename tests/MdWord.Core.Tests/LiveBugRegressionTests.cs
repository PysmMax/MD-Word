using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MdWord.Core;
using Xunit;

namespace MdWord.Core.Tests;

/// <summary>
/// Регресії на дефекти, знайдені користувачем у живому Word
/// (live tests/test 001.txt). До відповідних фіксів Стадії 1 ці тести падають.
/// </summary>
public class LiveBugRegressionTests
{
    // --- helpers ---------------------------------------------------------

    private static Body GetOoxmlBody(string markdown)
    {
        var result = new MarkdownConverter(null).ToOoxml(markdown);
        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.Document.Body!;
    }

    private static string BuildFlatOpc(params OpenXmlElement[] bodyChildren)
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body(bodyChildren));
        main.Document.Save();   // mirror the proven BuildFlatOpc pattern in OoxmlToMdTests
        return doc.ToFlatOpcString();
    }

    private static string ToMd(params OpenXmlElement[] bodyChildren) =>
        new MarkdownConverter(null).ToMarkdown(BuildFlatOpc(bodyChildren)).Markdown;

    // --- LIVE-1: <br> у клітинці на ВСТАВЦІ стає розривом, не текстом -----

    [Fact]
    public void Live1_Paste_BrInsideTableCell_BecomesLineBreak_NotLiteralText()
    {
        var body = GetOoxmlBody("| a | b |\n| --- | --- |\n| x<br>y | z |");
        var cell = body.Descendants<TableCell>().First(c => c.InnerText.Contains("x"));
        Assert.DoesNotContain("<br>", cell.InnerText);
        Assert.Single(cell.Descendants<Break>());
        Assert.Contains("x", cell.InnerText);
        Assert.Contains("y", cell.InnerText);
    }

    // --- LIVE-2: [1] при копіюванні не перетворюється на \[1\] (→ формула) -

    [Fact]
    public void Live2_Copy_SquareBracketCitation_IsNotEscapedIntoMathDelimiters()
    {
        var flatOpc = new MarkdownConverter(null).ToOoxml("див. [1] тут").FlatOpc;
        var md = new MarkdownConverter(null).ToMarkdown(flatOpc).Markdown;
        Assert.Contains("[1]", md);
        Assert.DoesNotContain("\\[", md);
    }

    // --- LIVE-3: розрив сторінки не лишає зайвий "\" --------------------

    [Fact]
    public void Live3_Copy_PageBreak_DoesNotLeaveStrayBackslash()
    {
        var md = ToMd(
            new Paragraph(new Run(new Text("сторінка 1"))),
            new Paragraph(new Run(new Break { Type = BreakValues.Page })),
            new Paragraph(new Run(new Text("сторінка 2"))));
        Assert.DoesNotContain("\\\n", md);
        Assert.Contains("сторінка 1", md);
        Assert.Contains("сторінка 2", md);
    }

    // --- LIVE-4: жирний рядок із хвостовим пробілом дає валідний MD ------

    [Fact]
    public void Live4_Copy_BoldRunWithTrailingSpace_ProducesValidBold()
    {
        var md = ToMd(new Paragraph(
            new Run(new RunProperties(new Bold()),
                    new Text("Система ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new Text("– текст"))));
        Assert.Contains("**Система**", md);
        Assert.DoesNotContain("**Система **", md);
    }

    // --- LIVE-5: символ шрифту Symbol (δ) не зникає ---------------------

    [Fact]
    public void Live5_Copy_SymbolFontGreekLetter_IsMapped_NotDropped()
    {
        var md = ToMd(new Paragraph(new Run(
            new SymbolChar { Font = "Symbol", Char = "F064" })));   // F064 = delta
        Assert.Contains("δ", md);
    }
}
