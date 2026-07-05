using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using MathOfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;

namespace MdWord.Core.Tests;

/// <summary>
/// Phase 4: OMML → MathML (Office's OMML2MML.XSL) → LaTeX (vendored xsltml
/// 2.1.2). Uses the real local Office XSL paths (same convention as
/// <see cref="MathConversionTests"/>) so these exercise the whole conveyor
/// for real. Builds real OMML via the forward direction
/// (<c>MarkdownConverter.ToOoxml</c>) first, then feeds the resulting
/// document back through <c>ToMarkdown</c> — this is closer to what actually
/// happens in the add-in than hand-writing OMML XML, and confirms the two
/// directions agree on shape.
///
/// Per the Phase 4 brief: assert the produced LaTeX contains the expected
/// command/structure, not exact string equality (a LaTeX→OMML→MathML→LaTeX
/// chain can legitimately reformat a formula, e.g. \dfrac vs \frac).
/// </summary>
public class MathReverseConversionTests
{
    private static readonly MathXslPaths XslPaths = OfficeXslLocator.Resolve();

    private static MarkdownResult RoundTripMath(string markdown)
    {
        var converter = new MarkdownConverter(XslPaths);
        var forward = converter.ToOoxml(markdown);
        return converter.ToMarkdown(forward.FlatOpc);
    }

    [Fact]
    public void InlineFormula_RoundTrips_AsDollarDelimitedLatex()
    {
        var reverse = RoundTripMath("Text with $x^2$ inline.");

        Assert.Contains("$", reverse.Markdown);
        Assert.DoesNotContain("`[formula]`", reverse.Markdown);
        Assert.Empty(reverse.Warnings);
    }

    [Fact]
    public void DisplayFormula_RoundTrips_AsDoubleDollarBlock()
    {
        var reverse = RoundTripMath("$$\nE=mc^2\n$$");

        Assert.Contains("$$", reverse.Markdown);
        Assert.DoesNotContain("`[formula]`", reverse.Markdown);
    }

    [Fact]
    public void Fraction_ProducesLatexContainingFracCommand()
    {
        var reverse = RoundTripMath("$\\frac{a}{b}$");

        Assert.Contains("\\frac", reverse.Markdown);
        Assert.Contains("{a}", reverse.Markdown);
        Assert.Contains("{b}", reverse.Markdown);
    }

    [Fact]
    public void SuperscriptAndSubscript_ProduceCaretAndUnderscore()
    {
        var reverse = RoundTripMath("$x^2$ and $x_i$");

        Assert.Contains("^", reverse.Markdown);
        Assert.Contains("_", reverse.Markdown);
    }

    [Fact]
    public void SquareRoot_ProducesSqrtCommand()
    {
        var reverse = RoundTripMath("$\\sqrt{x+1}$");

        Assert.Contains("\\sqrt", reverse.Markdown);
    }

    [Fact]
    public void SummationWithLimits_ProducesSumCommandWithLimits()
    {
        var reverse = RoundTripMath("$$\n\\sum_{i=1}^{n} i\n$$");

        Assert.Contains("\\sum", reverse.Markdown);
    }

    [Fact]
    public void Matrix_ProducesArrayEnvironmentWithParenDelimiters()
    {
        // mmltex (unlike KaTeX) has no special-case for the "pmatrix"
        // shorthand -- it reconstructs a parenthesized matrix as a generic
        // \left(\begin{array}{cc}...\end{array}\right), which is a legitimate
        // reformatting (same rendered result), not a bug. Confirmed
        // empirically before asserting on it, per the Phase 4 brief's own
        // warning against chasing LaTeX-spelling variance as a failure.
        var reverse = RoundTripMath("$$\n\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}\n$$");

        Assert.Contains("\\begin{array}", reverse.Markdown);
        Assert.Contains("\\left(", reverse.Markdown);
        Assert.Contains("\\right)", reverse.Markdown);
        Assert.Contains("1", reverse.Markdown);
        Assert.Contains("4", reverse.Markdown);
    }

    [Fact]
    public void NullOmml2MmlXslPath_DegradesToFormulaPlaceholder_WithWarning()
    {
        var xslPathsWithoutReverse = new MathXslPaths
        {
            Mml2OmmlXsl = XslPaths.Mml2OmmlXsl,
            Omml2MmlXsl = null,
        };

        var converter = new MarkdownConverter(xslPathsWithoutReverse);
        var forward = converter.ToOoxml("$x^2$");

        // Forward direction needs BOTH paths to activate OMML conversion
        // (MathConversionContext.IsActive) -- with Omml2MmlXsl missing, the
        // formula never became OMML in the first place, so it's still
        // literal "$x^2$" text; reading it back should just pass that
        // literal text through unchanged, no warning.
        var reverseFromLiteral = new MarkdownConverter(null).ToMarkdown(forward.FlatOpc);
        Assert.Contains("$x^2$", reverseFromLiteral.Markdown);
    }

    [Fact]
    public void MissingOmml2MmlXslPath_WhenRealOmmlPresent_DegradesToFormulaPlaceholder_WithWarning()
    {
        // Build real OMML using both paths, then read it back with only
        // Mml2OmmlXsl (no Omml2MmlXsl) -- this is the actual degrade path
        // this test targets: OMML exists, but the reverse conveyor can't run.
        var forward = new MarkdownConverter(XslPaths).ToOoxml("$x^2$");

        var xslPathsWithoutReverse = new MathXslPaths
        {
            Mml2OmmlXsl = XslPaths.Mml2OmmlXsl,
            Omml2MmlXsl = null,
        };

        var reverse = new MarkdownConverter(xslPathsWithoutReverse).ToMarkdown(forward.FlatOpc);

        Assert.Contains("`[formula]`", reverse.Markdown);
        Assert.NotEmpty(reverse.Warnings);
    }

    [Fact]
    public void MultiEquationMathPara_RoundTrips_BothFormulasPresent()
    {
        // Word can store 2+ equations in a single display block as one
        // m:oMathPara with multiple m:oMath children (REL-07). Build that
        // shape by hand -- splicing two single-equation OfficeMath elements,
        // each produced by the real forward pipeline, into one oMathPara --
        // and confirm both formulas survive the reverse direction instead of
        // one silently disappearing.
        var converter = new MarkdownConverter(XslPaths);
        var firstMath = ExtractOfficeMath(converter.ToOoxml("$$\nx^2\n$$").DocxBytes);
        var secondMath = ExtractOfficeMath(converter.ToOoxml("$$\ny_1\n$$").DocxBytes);

        var flatOpc = BuildFlatOpc(new Paragraph(new MathParagraph(firstMath, secondMath)));
        var reverse = converter.ToMarkdown(flatOpc);

        Assert.Contains("^", reverse.Markdown);
        Assert.Contains("_", reverse.Markdown);
        Assert.Equal(4, reverse.Markdown.Split(new[] { "$$" }, StringSplitOptions.None).Length - 1);
        Assert.Empty(reverse.Warnings);
    }

    private static MathOfficeMath ExtractOfficeMath(byte[] docxBytes)
    {
        using var stream = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var officeMath = doc.MainDocumentPart!.Document.Body!.Descendants<MathOfficeMath>().Single();
        return (MathOfficeMath)officeMath.CloneNode(true);
    }

    private static string BuildFlatOpc(params OpenXmlElement[] bodyChildren)
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body(bodyChildren));
        main.Document.Save();
        return doc.ToFlatOpcString();
    }
}
