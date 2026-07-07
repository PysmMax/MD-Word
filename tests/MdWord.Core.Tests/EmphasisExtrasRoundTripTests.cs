using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace MdWord.Core.Tests;

/// <summary>
/// Round trips for the v1.0.2 EmphasisExtras, in both directions, all with
/// xslPaths = null (no Office dependency — these must stay runnable on CI):
/// <list type="bullet">
/// <item>Markdown → OOXML → Markdown keeps the extras' delimiters.</item>
/// <item>Word text that merely LOOKS like the delimiters must come back as
/// the same literal text, not as formatting (escaping — otherwise pasting
/// the copied Markdown would silently reformat prose).</item>
/// </list>
/// </summary>
public class EmphasisExtrasRoundTripTests
{
    private static string RoundTripMarkdown(string markdown)
    {
        var converter = new MarkdownConverter(null);
        var forward = converter.ToOoxml(markdown);
        return converter.ToMarkdown(forward.FlatOpc).Markdown;
    }

    [Theory]
    [InlineData("H~2~O", "H~2~O")]
    [InlineData("x^2^", "x^2^")]
    [InlineData("==note==", "==note==")]
    [InlineData("++added++", "++added++")]
    [InlineData("~~gone~~", "~~gone~~")]
    public void MarkdownExtras_SurviveTheRoundTrip(string markdown, string expectedFragment)
    {
        var roundTripped = RoundTripMarkdown(markdown);

        Assert.Contains(expectedFragment, roundTripped);
    }

    private static string BuildFlatOpcWithLiteralText(string text)
    {
        using var stream = new MemoryStream();
        string flatOpc;

        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            mainPart.Document = new Document(new Body(new Paragraph(run)));
            mainPart.Document.Save();
            flatOpc = doc.ToFlatOpcString();
        }

        return flatOpc;
    }

    [Theory]
    [InlineData("approx ~10 to 20~ spread")]
    [InlineData("caret ^one and two^ here")]
    [InlineData("equals ==not a mark== here")]
    [InlineData("plus ++not inserted++ here")]
    public void LiteralWordText_ThatLooksLikeDelimiters_RoundTripsAsTheSameLiteralText(string text)
    {
        var converter = new MarkdownConverter(null);

        // Word → Markdown: the copy path must escape the would-be delimiters…
        var markdown = converter.ToMarkdown(BuildFlatOpcWithLiteralText(text)).Markdown;

        // …so that Markdown → Word reproduces the exact original text with
        // no formatting applied.
        var back = converter.ToOoxml(markdown);
        using var stream = new MemoryStream(back.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart.Document.Body;

        Assert.Equal(text, body.InnerText);
        Assert.Empty(body.Descendants<VerticalTextAlignment>());
        Assert.Empty(body.Descendants<Highlight>());
        Assert.Empty(body.Descendants<Underline>());
    }

    [Fact]
    public void LoneEqualsAndPlus_AreNotEscaped()
    {
        var converter = new MarkdownConverter(null);

        var markdown = converter.ToMarkdown(BuildFlatOpcWithLiteralText("a = b + c")).Markdown;

        Assert.Contains("a = b + c", markdown);
    }
}
