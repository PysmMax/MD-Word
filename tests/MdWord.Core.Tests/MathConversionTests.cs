using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OMath = DocumentFormat.OpenXml.Math;

namespace MdWord.Core.Tests;

/// <summary>
/// Phase 2: LaTeX → OMML. Uses the real local Office XSL paths (see
/// <see cref="XslPaths"/>) so these tests exercise KaTeX (via Jint) and
/// MML2OMML.XSL for real, not just the null-xslPaths degrade path already
/// covered by <see cref="InlineMappingTests.InlineMath_DegradesToLiteralTextWithDollarDelimiters"/>
/// and <see cref="BlockMappingTests.BlockMath_DegradesToParagraphWithDollarDollarDelimiters"/>.
///
/// Block math (<c>$$...$$</c>) needs the opening/closing <c>$$</c> on their
/// own lines to parse as Markdig.Extensions.Mathematics.MathBlock — a
/// single-line "$$E=mc^2$$" parses as inline display math instead, same as
/// BlockMappingTests' existing convention ("$$\nE=mc^2\n$$").
/// </summary>
[Trait("Category", "RequiresOfficeXsl")]
public class MathConversionTests
{
    private static readonly MathXslPaths XslPaths = OfficeXslLocator.Resolve();

    private static Body GetBody(string markdown, MathXslPaths xslPaths)
    {
        var converter = new MarkdownConverter(xslPaths);
        var result = converter.ToOoxml(markdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart.Document.Body;
    }

    private static Body GetBody(string markdown) => GetBody(markdown, XslPaths);

    private static (Body Body, string[] Warnings) GetBodyAndWarnings(string markdown, MathXslPaths xslPaths)
    {
        var converter = new MarkdownConverter(xslPaths);
        var result = converter.ToOoxml(markdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return (doc.MainDocumentPart.Document.Body, result.Warnings);
    }

    private static (Body Body, string[] Warnings) GetBodyAndWarnings(string markdown) => GetBodyAndWarnings(markdown, XslPaths);

    private static void AssertValid(string markdown, MathXslPaths xslPaths)
    {
        var converter = new MarkdownConverter(xslPaths);
        var result = converter.ToOoxml(markdown);

        using var stream = new MemoryStream(result.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator();
        var errors = validator.Validate(doc).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => e.Description)));
    }

    private static void AssertValid(string markdown) => AssertValid(markdown, XslPaths);

    [Fact]
    public void InlineMath_ConvertsToOfficeMath_EmbeddedInParagraphRunSequence()
    {
        var body = GetBody("before $E=mc^2$ after");

        var paragraph = body.Elements<Paragraph>().Single();
        var officeMath = paragraph.Descendants<OMath.OfficeMath>().Single();

        // Embedded inline, not its own paragraph: the surrounding text runs are
        // siblings of the OfficeMath element inside the same paragraph.
        Assert.Contains("before", paragraph.InnerText);
        Assert.Contains("after", paragraph.InnerText);
        Assert.Null(paragraph.Descendants<OMath.Paragraph>().SingleOrDefault());
        Assert.NotEmpty(officeMath.Descendants<OMath.Run>());

        AssertValid("before $E=mc^2$ after");
    }

    [Fact]
    public void DisplayMath_ConvertsToOfficeMathParagraph_AsItsOwnParagraph()
    {
        var body = GetBody("$$\nE=mc^2\n$$");

        var paragraph = body.Elements<Paragraph>().Single();
        var oMathPara = paragraph.Elements<OMath.Paragraph>().Single();
        var officeMath = oMathPara.Elements<OMath.OfficeMath>().Single();

        Assert.NotEmpty(officeMath.Descendants<OMath.Run>());

        AssertValid("$$\nE=mc^2\n$$");
    }

    [Fact]
    public void Fraction_ProducesMFractionElement()
    {
        var body = GetBody(@"$\frac{a}{b}$");

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var fraction = officeMath.Descendants<OMath.Fraction>().Single();

        Assert.Contains("a", fraction.Descendants<OMath.Numerator>().Single().InnerText);
        Assert.Contains("b", fraction.Descendants<OMath.Denominator>().Single().InnerText);
    }

    [Fact]
    public void SquareRoot_ProducesMRadElement()
    {
        var body = GetBody(@"$\sqrt{x+1}$");

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var radical = officeMath.Descendants<OMath.Radical>().Single();

        Assert.Contains("x", radical.InnerText);
        Assert.Contains("1", radical.InnerText);
    }

    [Fact]
    public void SubscriptSuperscript_ProducesMSSubSupElement()
    {
        var body = GetBody(@"$x_i^2$");

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var subSup = officeMath.Descendants<OMath.SubSuperscript>().Single();

        Assert.Contains("i", subSup.Elements<OMath.SubArgument>().Single().InnerText);
        Assert.Contains("2", subSup.Elements<OMath.SuperArgument>().Single().InnerText);
    }

    [Fact]
    public void Summation_ProducesMNaryElement()
    {
        var body = GetBody(@"$\sum_{i=1}^{n}$");

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var nary = officeMath.Descendants<OMath.Nary>().Single();

        Assert.NotNull(nary);
    }

    [Fact]
    public void Integral_AbsorbsIntegrandIntoNaryArgument_NotAsOrphanedSibling()
    {
        var markdown = "$$\n\\int_{-\\infty}^{\\infty} e^{-x^2}\\,dx = \\sqrt{\\pi}\n$$";
        var body = GetBody(markdown);

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var nary = officeMath.Descendants<OMath.Nary>().Single();
        var naryBase = nary.Elements<OMath.Base>().Single();

        // The integrand ("e^{-x^2} dx") must be nested inside the nary's own
        // argument slot (<m:e>), not left as orphaned siblings after
        // </m:nary> -- MML2OMML.XSL's NaryHandleMrowMstyle only pulls in an
        // argument when the single following sibling is an mrow/mstyle;
        // KaTeX emits flat siblings, so without pre-grouping this slot stays
        // empty and Word renders "e^{-x^2} dx" as text disconnected from a
        // dashed empty-placeholder box.
        Assert.Contains("e", naryBase.InnerText);
        Assert.Contains("dx", naryBase.InnerText);

        // "= √π" is a separate relation, not part of the integral -- it must
        // stay outside the nary's argument.
        Assert.DoesNotContain("=", naryBase.InnerText);
        Assert.DoesNotContain("π", naryBase.InnerText);

        AssertValid(markdown);
    }

    [Fact]
    public void IndefiniteIntegral_WithNoLimits_ProducesMNaryElement_WithHiddenLimits()
    {
        var markdown = "$$\n\\int f(x)\\,dx = F(x) + C\n$$";
        var body = GetBody(markdown);

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var nary = officeMath.Descendants<OMath.Nary>().Single();
        var naryBase = nary.Elements<OMath.Base>().Single();

        // MML2OMML.XSL has no template at all for a bare, unwrapped nary
        // operator (confirmed: no `match="mml:mo"` template in the XSL) --
        // KaTeX emits \int with no {}^{} limits as a plain sibling <mo>,
        // which the transform falls back to serializing as flat literal
        // text ("∫f(x)dx"), no <m:nary> at all. Since MML2OMML.XSL can never
        // produce subHide=on + supHide=on together (its own limLoc/hide
        // table only reachable via msub/msup/msubsup/munder/mover/munderover,
        // none of which yields "both hidden"), NaryArgumentGrouper
        // synthesizes a wrapper and a later pass hides the resulting empty
        // limit slots.
        Assert.Contains("f", naryBase.InnerText);
        Assert.Contains("dx", naryBase.InnerText);
        Assert.DoesNotContain("=", naryBase.InnerText);

        var naryPr = nary.GetFirstChild<OMath.NaryProperties>();
        Assert.Contains("subHide", naryPr.OuterXml);
        Assert.Contains("supHide", naryPr.OuterXml);
        Assert.DoesNotContain("m:subHide m:val=\"off\"", naryPr.OuterXml);
        Assert.DoesNotContain("m:supHide m:val=\"off\"", naryPr.OuterXml);

        AssertValid(markdown);
    }

    [Fact]
    public void BareSummation_WithNoLimits_ProducesMNaryElement_WithHiddenLimits()
    {
        // Same underlying gap as the indefinite-integral case, for a
        // different n-ary character -- confirms the fix is general, not
        // integral-specific.
        var body = GetBody(@"$\sum x_i$");

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var nary = officeMath.Descendants<OMath.Nary>().Single();
        var naryBase = nary.Elements<OMath.Base>().Single();

        Assert.Contains("x", naryBase.InnerText);
        Assert.Contains("i", naryBase.InnerText);
    }

    [Theory]
    [InlineData(@"A \cap B")]
    [InlineData(@"A \cup B")]
    public void BinarySetOperator_WithNoLimits_StaysFlat_NotWrappedAsNary(string tex)
    {
        // \cap/\cup (∩/∪, U+2229/U+222A) are binary infix operators when
        // used bare -- unlike \int/\sum/\bigcap, wrapping a bare one as a
        // prefix n-ary would orphan its left operand and swallow its right
        // operand as the n-ary's sole argument. Only the "big" variants
        // (\bigcap/\bigcup -> ⋂/⋃, U+22C2/U+22C3) are genuinely prefix and
        // safe to bare-wrap.
        var body = GetBody("$" + tex + "$");

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        Assert.Empty(officeMath.Descendants<OMath.Nary>());
        Assert.Contains("A", officeMath.InnerText);
        Assert.Contains("B", officeMath.InnerText);
    }

    [Fact]
    public void Summation_WithSummand_AbsorbsSummandIntoNaryArgument()
    {
        var body = GetBody(@"$\sum_{i=1}^{n} i$");

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var nary = officeMath.Descendants<OMath.Nary>().Single();
        var naryBase = nary.Elements<OMath.Base>().Single();

        Assert.Equal("i", naryBase.InnerText);
    }

    [Fact]
    public void GreekLetters_RenderAsTheirUnicodeCharacters()
    {
        var body = GetBody(@"$\alpha + \pi$");

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var text = officeMath.InnerText;

        Assert.Contains("α", text); // \alpha
        Assert.Contains("π", text); // \pi
    }

    [Fact]
    public void Matrix_ProducesMMatrixElement_WithExpectedRowsAndColumns()
    {
        var body = GetBody("$$\n\\begin{pmatrix} 1 & 2 \\\\ 3 & 4 \\end{pmatrix}\n$$");

        var officeMath = body.Descendants<OMath.OfficeMath>().Single();
        var matrix = officeMath.Descendants<OMath.Matrix>().Single();
        var rows = matrix.Elements<OMath.MatrixRow>().ToList();

        Assert.Equal(2, rows.Count);
        Assert.All(rows, row => Assert.Equal(2, row.Elements<OMath.Base>().Count()));
        Assert.Equal("1", rows[0].Elements<OMath.Base>().First().InnerText);
        Assert.Equal("4", rows[1].Elements<OMath.Base>().Last().InnerText);
    }

    [Fact]
    public void BrokenFormula_DegradesToLiteralText_AndAddsWarning_AndDocumentStaysValid()
    {
        var markdown = @"before $\frac{a}{$ after";
        var (body, warnings) = GetBodyAndWarnings(markdown);

        var paragraph = body.Elements<Paragraph>().Single();
        Assert.Contains(paragraph.Elements<Run>(), r => r.InnerText.Contains(@"\frac{a}{"));
        Assert.Empty(paragraph.Descendants<OMath.OfficeMath>());

        Assert.NotEmpty(warnings);
        Assert.Contains(warnings, w => w.Contains(@"\frac{a}{"));

        AssertValid(markdown);
    }

    [Fact]
    public void ParenAndBracketDelimiters_ProduceSameOmmlAsDollarForm()
    {
        var dollarBody = GetBody(@"$E=mc^2$");
        var parenBody = GetBody(@"\(E=mc^2\)");

        var dollarMath = dollarBody.Descendants<OMath.OfficeMath>().Single();
        var parenMath = parenBody.Descendants<OMath.OfficeMath>().Single();
        Assert.Equal(dollarMath.OuterXml, parenMath.OuterXml);

        // The preprocessor is deliberately line-oriented (the initial plan, Phase 2 task
        // list) -- \[...\] only rewrites within a single line, so this checks
        // the single-line form; both "$$E=mc^2$$" and its rewritten
        // "\[E=mc^2\]" source parse as inline display math (see the
        // DisplayMath_* test's comment on why $$...$$ needs its own lines to
        // become a MathBlock), which is fine for an equivalence check.
        var dollarBlockBody = GetBody("$$E=mc^2$$");
        var bracketBlockBody = GetBody("\\[E=mc^2\\]");

        var dollarBlockMath = dollarBlockBody.Descendants<OMath.OfficeMath>().Single();
        var bracketBlockMath = bracketBlockBody.Descendants<OMath.OfficeMath>().Single();
        Assert.Equal(dollarBlockMath.OuterXml, bracketBlockMath.OuterXml);
    }

    [Fact]
    public void ParenDelimiters_InsideFencedCodeBlock_AreLeftUntouched()
    {
        var body = GetBody("```\nsee \\(E=mc^2\\) literally\n```");

        var paragraph = body.Elements<Paragraph>().Single();
        Assert.Empty(paragraph.Descendants<OMath.OfficeMath>());
        Assert.Equal(@"see \(E=mc^2\) literally", paragraph.InnerText);
    }

    [Fact]
    public void NullXslPaths_KeepsLiteralTextDegradation_WithoutWarnings()
    {
        var (body, warnings) = GetBodyAndWarnings("$E=mc^2$", null);

        var paragraph = body.Elements<Paragraph>().Single();
        Assert.Contains(paragraph.Elements<Run>(), r => r.InnerText == "$E=mc^2$");
        Assert.Empty(paragraph.Descendants<OMath.OfficeMath>());
        Assert.Empty(warnings);
    }
}
