using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace MdWord.Core.Tests;

/// <summary>
/// Regression pack for the 2026-07 code review findings (REL-01..REL-06):
/// inputs that previously lost content silently, aborted the whole
/// conversion, or produced wrong formatting. Each fact names its finding.
/// </summary>
public class RegressionTests
{
    private static (Body Body, string[] Warnings) GetBodyAndWarnings(string markdown)
    {
        var converter = new MarkdownConverter(null);
        var result = converter.ToOoxml(markdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return (doc.MainDocumentPart.Document.Body, result.Warnings);
    }

    [Fact] // REL-01
    public void TableInsideQuote_IsNotSilentlyDropped()
    {
        var (body, _) = GetBodyAndWarnings("> | a | b |\n> | --- | --- |\n> | 1 | 2 |");

        var table = body.Elements<Table>().Single();
        Assert.Contains("a", table.InnerText);
        Assert.Contains("2", table.InnerText);
    }

    [Fact] // REL-01
    public void CodeBlockInsideListItem_IsNotSilentlyDropped()
    {
        var (body, _) = GetBodyAndWarnings("- item\n\n  ```\n  code inside item\n  ```");

        Assert.Contains("code inside item", body.InnerText);
        Assert.Contains("item", body.InnerText);
    }

    [Fact] // REL-02
    public void LinkWithSpaceInAbsoluteUrl_DegradesInsteadOfThrowing()
    {
        var (body, warnings) = GetBodyAndWarnings("[t](<http://exa mple.com>)");

        Assert.Contains("t", body.InnerText);
        Assert.NotEmpty(warnings);
    }

    // v1.0.2 intentionally reverses the REL-04 literal-text degradation:
    // EmphasisExtras.Default is now enabled, so these delimiters map to real
    // OOXML formatting instead of being left as literal source text.
    [Fact]
    public void SingleTildeSubscript_MapsToSubscript_NotStrikethrough()
    {
        var (body, _) = GetBodyAndWarnings("H~2~O");

        var paragraph = body.Elements<Paragraph>().Single();
        Assert.Equal("H2O", paragraph.InnerText);
        Assert.Empty(paragraph.Descendants<Strike>());
        // Single tilde is subscript, not strikethrough
        var subscriptRun = paragraph.Elements<Run>().Single(r => r.InnerText == "2");
        Assert.NotNull(subscriptRun.RunProperties?.VerticalTextAlignment);
    }

    [Fact]
    public void CaretSuperscriptAndMarkedText_MapToSuperscriptAndHighlight()
    {
        var (body, _) = GetBodyAndWarnings("x^2^ and ==marked==");

        var paragraph = body.Elements<Paragraph>().Single();
        Assert.Equal("x2 and marked", paragraph.InnerText);
        Assert.Empty(paragraph.Descendants<Italic>());
        Assert.Empty(paragraph.Descendants<Bold>());
        // Caret is superscript and equals is highlight, not italic/bold
        var superscriptRun = paragraph.Elements<Run>().Single(r => r.InnerText == "2");
        Assert.NotNull(superscriptRun.RunProperties?.VerticalTextAlignment);
        var markedRun = paragraph.Elements<Run>().Single(r => r.InnerText == "marked");
        Assert.NotNull(markedRun.RunProperties?.Highlight);
    }

    [Fact] // REL-06
    public void Image_DegradesToAltTextWithWarning_NotHyperlink()
    {
        var (body, warnings) = GetBodyAndWarnings("![alt text](https://example.com/img.png)");

        Assert.Empty(body.Descendants<Hyperlink>());
        Assert.Contains("alt text", body.InnerText);
        Assert.NotEmpty(warnings);
    }

    [Fact] // REL-03
    public void MultiLineBracketDelimiters_BecomeDisplayMathBlock()
    {
        // Without Office XSL paths math degrades to the literal "$$...$$"
        // paragraph -- same shape BlockMappingTests already asserts for
        // "$$\nE=mc^2\n$$". The point: "\[" / "\]" on their own lines must
        // reach the SAME path, not be eaten as escaped brackets.
        var (body, _) = GetBodyAndWarnings("\\[\nE=mc^2\n\\]");

        var paragraph = body.Elements<Paragraph>().Single();
        Assert.Equal("$$E=mc^2$$", paragraph.InnerText);
    }

    [Fact] // REL-05
    public void HardBreakInsideTableCell_RendersAsBrTag_ReverseDirection()
    {
        var converter = new MarkdownConverter(null);
        var forward = converter.ToOoxml("| a | b |\n| --- | --- |\n| 1 | 2 |");

        // Inject a <w:br/> into the first data cell of the generated table.
        using var stream = new MemoryStream();
        stream.Write(forward.DocxBytes, 0, forward.DocxBytes.Length);
        stream.Position = 0;
        string flatOpc;
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var run = doc.MainDocumentPart.Document.Body
                .Descendants<TableCell>()
                .First(c => c.InnerText == "1")
                .Descendants<Run>()
                .First();
            run.Append(new Break());
            run.Append(new Text("next"));
            doc.MainDocumentPart.Document.Save();
            flatOpc = doc.ToFlatOpcString();
        }

        var reverse = converter.ToMarkdown(flatOpc);
        var tableLines = reverse.Markdown.Split('\n').Where(l => l.TrimStart().StartsWith("|")).ToList();

        Assert.Contains(tableLines, l => l.Contains("1<br>next"));
        Assert.DoesNotContain("\\\n", string.Join("\n", tableLines));
    }
}
