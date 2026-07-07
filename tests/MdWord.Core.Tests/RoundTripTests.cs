using System;
using System.IO;
using System.Text.RegularExpressions;
using Xunit;

namespace MdWord.Core.Tests;

/// <summary>
/// The real correctness bar for Phase 4 (per the brief): <c>demo.md</c> →
/// <c>ToOoxml</c> → <c>ToMarkdown</c> should be structurally equivalent to
/// <c>demo.md</c> — same heading text at the same levels, same table cell
/// contents, same list items — not byte-identical (Markdown has multiple
/// valid spellings of the same structure). Exercises the forward and
/// reverse walkers against each other without needing live Word.
///
/// Uses the real local Office XSL paths (same convention as
/// <see cref="MathConversionTests"/>/<see cref="MathReverseConversionTests"/>)
/// so the math elements go through the real OMML round-trip too, not the
/// null-xslPaths literal-text degrade path — a stronger integration check.
/// Math assertions stay loose per the brief: just that each formula
/// survives with non-empty content, no LaTeX-string-equality assertions.
/// </summary>
[Trait("Category", "RequiresOfficeXsl")]
public class RoundTripTests
{
    private static readonly MathXslPaths XslPaths = OfficeXslLocator.Resolve();

    private static string DemoMarkdown =>
        File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "samples", "demo.md"));

    private static string RoundTrip()
    {
        var converter = new MarkdownConverter(XslPaths);
        var forward = converter.ToOoxml(DemoMarkdown);
        var reverse = converter.ToMarkdown(forward.FlatOpc);
        return reverse.Markdown;
    }

    [Fact]
    public void Demo_RoundTrip_PreservesAllSixHeadingLevelsAndText()
    {
        var markdown = RoundTrip();

        Assert.Contains("# Heading H1", markdown);
        Assert.Contains("## Heading H2", markdown);
        Assert.Contains("### Heading H3", markdown);
        Assert.Contains("#### Heading H4", markdown);
        Assert.Contains("##### Heading H5", markdown);
        Assert.Contains("###### Heading H6", markdown);
    }

    [Fact]
    public void Demo_RoundTrip_PreservesTableCellContents()
    {
        var markdown = RoundTrip();

        Assert.Contains("Column A", markdown);
        Assert.Contains("Column B", markdown);
        Assert.Contains("Column C", markdown);
        Assert.Contains("regular", markdown);
        Assert.Contains("one", markdown);
        Assert.Contains("two", markdown);
        Assert.Contains("three", markdown);
    }

    [Fact]
    public void Demo_RoundTrip_PreservesListItemText()
    {
        var markdown = RoundTrip();

        Assert.Contains("First list item", markdown);
        Assert.Contains("Second list item", markdown);
        Assert.Contains("Third list item", markdown);
        Assert.Contains("Nested ordered one", markdown);
        Assert.Contains("Nested ordered two", markdown);
        Assert.Contains("Nested bullet inside the ordered list", markdown);
        Assert.Contains("First ordered item", markdown);
        Assert.Contains("Second ordered item", markdown);
        Assert.Contains("Third ordered item", markdown);
    }

    [Fact]
    public void Demo_RoundTrip_PreservesCodeContent()
    {
        var markdown = RoundTrip();

        Assert.Contains("line one of code", markdown);
        Assert.Contains("line two of code", markdown);
        Assert.Contains("indented code line", markdown);
        Assert.Contains("inline code", markdown);
    }

    [Fact]
    public void Demo_RoundTrip_PreservesQuote()
    {
        var markdown = RoundTrip();

        Assert.Contains("A quote with", markdown);
        Assert.Contains("Second paragraph of the quote", markdown);
    }

    [Fact]
    public void Demo_RoundTrip_PreservesCyrillicAndEmoji()
    {
        var markdown = RoundTrip();

        Assert.Contains("Привіт, світ!", markdown);
        Assert.Contains("🚀", markdown);
    }

    [Fact]
    public void Demo_RoundTrip_BothFormulasSurviveWithNonEmptyContent()
    {
        var markdown = RoundTrip();

        // Loose per the brief: non-empty $...$ / $$...$$ survives, no
        // LaTeX-string-equality assertion (mmltex may legitimately reformat
        // e.g. \dfrac vs \frac without being wrong).
        var inlineMatch = Regex.Match(markdown, @"\$([^$\n]+)\$");
        Assert.True(inlineMatch.Success, "Expected an inline $...$ formula in the result.");
        Assert.False(string.IsNullOrWhiteSpace(inlineMatch.Groups[1].Value));

        var displayMatch = Regex.Match(markdown, @"\$\$\s*([\s\S]+?)\s*\$\$");
        Assert.True(displayMatch.Success, "Expected a display $$...$$ formula in the result.");
        Assert.False(string.IsNullOrWhiteSpace(displayMatch.Groups[1].Value));
    }

    [Fact]
    public void Demo_RoundTrip_ContainsThematicBreak()
    {
        var markdown = RoundTrip();

        Assert.Contains("---", markdown);
    }

    [Fact]
    public void Demo_RoundTrip_HasNoUnexpectedWarnings()
    {
        var converter = new MarkdownConverter(XslPaths);
        var forward = converter.ToOoxml(DemoMarkdown);
        var reverse = converter.ToMarkdown(forward.FlatOpc);

        // demo.md has no images/footnotes/content-controls/merged cells, so
        // the reverse walk over our own generated output should be
        // warning-free -- a non-empty Warnings here would mean either a
        // regression or a demo.md fixture element this walker doesn't
        // recognize yet.
        Assert.Empty(reverse.Warnings);
    }
}
