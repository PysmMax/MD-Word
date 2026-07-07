using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace MdWord.Core.Tests;

/// <summary>
/// Reverse mapping (OOXML → Markdown) for the v1.0.2 EmphasisExtras:
/// w:vertAlign subscript/superscript → ~x~ / ^x^, w:highlight (any color,
/// not just yellow — foreign documents) → ==x==, w:u (any non-none value)
/// → ++x++. Hand-built documents via the OpenXML object model, mirroring
/// <see cref="OoxmlToMdTests"/>' foreign-document fixture style.
/// </summary>
public class EmphasisExtrasReverseTests
{
    private static string BuildFlatOpc(Body body)
    {
        using var stream = new MemoryStream();
        string flatOpc;

        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(body);
            mainPart.Document.Save();
            flatOpc = doc.ToFlatOpcString();
        }

        return flatOpc;
    }

    private static string ToMd(params OpenXmlElement[] paragraphChildren) =>
        new MarkdownConverter(null)
            .ToMarkdown(BuildFlatOpc(new Body(new Paragraph(paragraphChildren))))
            .Markdown;

    private static Run RunWith(RunProperties properties, string text)
    {
        var run = new Run();
        if (properties != null)
        {
            run.RunProperties = properties;
        }

        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    [Fact]
    public void SubscriptRun_EmitsSingleTildeMarkers()
    {
        var markdown = ToMd(
            RunWith(null, "H"),
            RunWith(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }), "2"),
            RunWith(null, "O"));

        Assert.Contains("H~2~O", markdown);
    }

    [Fact]
    public void SuperscriptRun_EmitsCaretMarkers()
    {
        var markdown = ToMd(
            RunWith(null, "x"),
            RunWith(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }), "2"));

        Assert.Contains("x^2^", markdown);
    }

    [Fact]
    public void BaselineVertAlign_EmitsNoMarkers()
    {
        var markdown = ToMd(
            RunWith(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Baseline }), "plain"));

        Assert.Contains("plain", markdown);
        Assert.DoesNotContain("~plain~", markdown);
        Assert.DoesNotContain("^plain^", markdown);
    }

    [Fact]
    public void HighlightedRun_AnyColor_EmitsDoubleEquals()
    {
        var markdown = ToMd(
            RunWith(new RunProperties(new Highlight { Val = HighlightColorValues.Green }), "note"));

        Assert.Contains("==note==", markdown);
    }

    [Fact]
    public void UnderlinedRun_EmitsDoublePlus()
    {
        var markdown = ToMd(
            RunWith(new RunProperties(new Underline { Val = UnderlineValues.Single }), "added"));

        Assert.Contains("++added++", markdown);
    }

    [Fact]
    public void UnderlineValueNone_EmitsNoMarkers()
    {
        var markdown = ToMd(
            RunWith(new RunProperties(new Underline { Val = UnderlineValues.None }), "plain"));

        Assert.Contains("plain", markdown);
        Assert.DoesNotContain("++", markdown);
    }

    [Fact]
    public void AdjacentRunsWithSameSubscript_AggregateIntoOneMarkerPair()
    {
        var markdown = ToMd(
            RunWith(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }), "a"),
            RunWith(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }), "b"));

        Assert.Contains("~ab~", markdown);
    }

    [Fact]
    public void StrikeAndSubscriptTogether_EmitOnlyStrike_NeverTripleTilde()
    {
        // '~~~x~~~' does not parse at all in Markdig (verified empirically),
        // so the combined case degrades to strikethrough only.
        var markdown = ToMd(
            RunWith(
                new RunProperties(
                    new Strike(),
                    new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }),
                "x"));

        Assert.Contains("~~x~~", markdown);
        Assert.DoesNotContain("~~~", markdown);
    }

    [Fact]
    public void BoldHighlightedRun_NestsBoldInsideEquals()
    {
        var markdown = ToMd(
            RunWith(new RunProperties(new Bold(), new Highlight { Val = HighlightColorValues.Yellow }), "hot"));

        Assert.Contains("==**hot**==", markdown);
    }
}
