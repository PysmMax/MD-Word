using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace MdWord.Core.Tests;

/// <summary>
/// Phase 4 reverse mapping (OOXML → Markdown) — one fixture per mapping,
/// mirroring <see cref="BlockMappingTests"/>/<see cref="InlineMappingTests"/>'
/// forward-direction coverage in reverse. Two fixture styles, per the brief:
/// <list type="bullet">
/// <item>Round-tripping our own <c>ToOoxml</c> output — exercises the
/// "own documents" path (styleId ∈ HeadingN, numId ∈ {1,2}, ...).</item>
/// <item>Hand-built documents via the OpenXML object model (not
/// <c>ToOoxml</c>) — exercises the "foreign document" path (renamed heading
/// styles, arbitrary numId, raw Consolas runs with no CodeInline style) that
/// round-tripping our own output can never reach, since our own generator
/// always uses the same styleIds/numIds. Building via the object model
/// rather than hand-editing XML strings avoids OOXML-shape mistakes while
/// still producing a document our own generator never would.</item>
/// </list>
/// </summary>
public class OoxmlToMdTests
{
    private static MarkdownResult ToMarkdown(string markdown)
    {
        var converter = new MarkdownConverter(null);
        var forward = converter.ToOoxml(markdown);
        return converter.ToMarkdown(forward.FlatOpc);
    }

    private static string BuildFlatOpc(Body body, Styles styles = null, Numbering numbering = null)
    {
        using var stream = new MemoryStream();
        string flatOpc;

        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(body);

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = styles ?? new Styles();
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

    // ---- Heading: three-way fallback ----

    [Theory]
    [InlineData("# H1", 1)]
    [InlineData("###### H6", 6)]
    public void Heading_OwnStyleId_RoundTrips(string markdown, int level)
    {
        var reverse = ToMarkdown(markdown);

        Assert.Contains(new string('#', level) + " ", reverse.Markdown);
    }

    [Fact]
    public void Heading_ForeignStyleId_ResolvesViaStyleName()
    {
        // A foreign document might rename/duplicate the heading style
        // (e.g. localized "Заголовок 2") while keeping Word's invariant
        // w:name "heading 2" -- must still resolve to H2 without relying on
        // the styleId literally being "Heading2".
        var body = new Body(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Zagholovok2" }),
                new Run(new Text("Foreign heading"))));

        var styles = new Styles(
            new Style(new StyleName { Val = "heading 2" }) { Type = StyleValues.Paragraph, StyleId = "Zagholovok2" });

        var flatOpc = BuildFlatOpc(body, styles);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("## Foreign heading", reverse.Markdown);
    }

    [Fact]
    public void Heading_OutlineLevelOnly_ResolvesWithoutAnyHeadingStyle()
    {
        // outlineLvl is 0-based; level 2 (0-based) => H3.
        var body = new Body(
            new Paragraph(
                new ParagraphProperties(new OutlineLevel { Val = 2 }),
                new Run(new Text("Outline-only heading"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("### Outline-only heading", reverse.Markdown);
    }

    [Fact]
    public void HeadingText_LeadingDash_IsEscaped()
    {
        // Bug 3c: EscapeLeadingMarker is applied at the start of heading
        // text too, not just plain paragraphs.
        var body = new Body(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new Run(new Text("- not a list"))));

        var styles = new Styles(
            new Style(new StyleName { Val = "heading 1" }) { Type = StyleValues.Paragraph, StyleId = "Heading1" });

        var flatOpc = BuildFlatOpc(body, styles);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("# \\- not a list", reverse.Markdown);
    }

    // ---- Emphasis: adjacent-run aggregation ----

    [Fact]
    public void AdjacentBoldRuns_AggregateIntoSingleDelimitedSpan()
    {
        var body = new Body(
            new Paragraph(
                new Run(new RunProperties(new Bold()), new Text("Hello ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new RunProperties(new Bold()), new Text("World"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("**Hello World**", reverse.Markdown);
        Assert.DoesNotContain("****", reverse.Markdown.Replace("**Hello World**", string.Empty));
    }

    [Fact]
    public void BoldItalic_AggregatesToTripleAsterisk()
    {
        var reverse = ToMarkdown("***bold italic***");

        Assert.Contains("***bold italic***", reverse.Markdown);
    }

    [Fact]
    public void StrikeAndBold_NestStrikeOutsideBold()
    {
        var body = new Body(
            new Paragraph(
                new Run(new RunProperties(new Bold(), new Strike()), new Text("both"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("~~**both**~~", reverse.Markdown);
    }

    // ---- Code: styleId AND raw monospace-font run (foreign document) ----

    [Fact]
    public void InlineCode_ViaCodeInlineStyle_RoundTrips()
    {
        var reverse = ToMarkdown("before `code` after");

        Assert.Contains("`code`", reverse.Markdown);
    }

    [Fact]
    public void InlineCode_ViaRawConsolasRun_NoStyleId_IsRecognizedAsCode()
    {
        // Foreign document: a plain Consolas run with no CodeInline styleId
        // at all, mixed into a paragraph with ordinary text either side.
        var body = new Body(
            new Paragraph(
                new Run(new Text("before ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(
                    new RunProperties(new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" }),
                    new Text("rawcode")),
                new Run(new Text(" after") { Space = SpaceProcessingModeValues.Preserve })));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("`rawcode`", reverse.Markdown);
        Assert.Contains("before", reverse.Markdown);
        Assert.Contains("after", reverse.Markdown);
    }

    [Fact]
    public void CodeBlock_ViaCodeBlockStyle_RoundTrips()
    {
        var reverse = ToMarkdown("```\nline one\nline two\n```");

        Assert.Contains("```", reverse.Markdown);
        Assert.Contains("line one", reverse.Markdown);
        Assert.Contains("line two", reverse.Markdown);
    }

    [Fact]
    public void CodeBlock_ViaWholeParagraphRawConsolas_NoCodeBlockStyle_IsRecognizedAsFencedCode()
    {
        // Foreign document: a whole paragraph entirely in Consolas, no
        // CodeBlock styleId anywhere.
        var body = new Body(
            new Paragraph(
                new Run(
                    new RunProperties(new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" }),
                    new Text("var x = 1;"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("```", reverse.Markdown);
        Assert.Contains("var x = 1;", reverse.Markdown);
    }

    [Fact]
    public void InlineCode_ContentWithDoubleBacktick_UsesTripleBacktickFence()
    {
        // CommonMark code-span rule: the fence must be one backtick longer
        // than the longest backtick run inside the content, or a fence that
        // matches a run inside the content would close the span early.
        var body = new Body(
            new Paragraph(
                new Run(new RunProperties(new RunStyle { Val = "CodeInline" }), new Text("a``b"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("```a``b```", reverse.Markdown);
    }

    [Fact]
    public void CodeBlock_ContentContainingTripleBacktick_UsesFourBacktickFence()
    {
        var body = new Body(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "CodeBlock" }),
                new Run(new Text("before ``` after"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("````\nbefore ``` after\n````", reverse.Markdown);
    }

    // ---- Hidden text (w:vanish) ----

    [Fact]
    public void HiddenRun_IsSkippedFromOutput_SurroundingTextSurvives()
    {
        var body = new Body(
            new Paragraph(
                new Run(new Text("Visible AAA ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new RunProperties(new Vanish()), new Text("HIDDENBBB")),
                new Run(new Text(" Visible CCC") { Space = SpaceProcessingModeValues.Preserve })));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("Visible AAA", reverse.Markdown);
        Assert.Contains("Visible CCC", reverse.Markdown);
        Assert.DoesNotContain("HIDDENBBB", reverse.Markdown);
    }

    [Fact]
    public void VanishExplicitFalse_IsNotTreatedAsHidden_TextSurvives()
    {
        // OOXML toggle-property semantics: <w:vanish w:val="false"/> means
        // NOT hidden -- the element's mere presence must not be read as
        // "hidden", same as Bold/Italic/Strike elsewhere in this file.
        var body = new Body(
            new Paragraph(
                new Run(new RunProperties(new Vanish { Val = false }), new Text("NotActuallyHidden"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("NotActuallyHidden", reverse.Markdown);
    }

    // ---- Tracked-change moves ----

    [Fact]
    public void MoveToRun_KeepsMovedTextInOutput()
    {
        // w:moveTo is the destination of a tracked-change move -- the moved
        // text now lives here as current document content, same as w:ins.
        var body = new Body(
            new Paragraph(
                new MoveToRun(new Run(new Text("MOVEDTEXT"))) { Id = "1" }));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("MOVEDTEXT", reverse.Markdown);
    }

    [Fact]
    public void MoveFromRun_IsStillSkipped_ControlCase()
    {
        // w:moveFrom is the old location text moved away from -- must stay
        // skipped (unchanged control case for the MoveToRun fix above).
        var body = new Body(
            new Paragraph(
                new MoveFromRun(new Run(new Text("OLDTEXT"))) { Id = "1" }));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.DoesNotContain("OLDTEXT", reverse.Markdown);
    }

    // ---- Quote ----

    [Fact]
    public void Quote_SingleParagraph_RoundTrips()
    {
        var reverse = ToMarkdown("> quoted text");

        Assert.Contains("> quoted text", reverse.Markdown);
    }

    [Fact]
    public void Quote_MultipleParagraphs_JoinsWithBareGtLineBetween()
    {
        var reverse = ToMarkdown("> first\n>\n> second");

        Assert.Contains("> first\n>\n> second", reverse.Markdown);
    }

    [Fact]
    public void QuoteText_LeadingGreaterThan_IsEscaped_NotReadAsNestedBlockquote()
    {
        // "> " immediately followed by another "> " is CommonMark syntax for
        // a nested blockquote -- a quote paragraph whose own text starts
        // with ">" must have that leading marker escaped too. In practice
        // this ">" is now caught by bug 3a's always-escape rule in
        // EscapeInlineText (">" backslash-escaped everywhere, not just at a
        // line start) before EscapeLeadingMarker(s) (bug 3c) even runs --
        // either mechanism getting it is what this test actually checks.
        var body = new Body(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Quote" }),
                new Run(new Text("> nested-looking"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("> \\> nested-looking", reverse.Markdown);
    }

    // ---- Hyperlinks ----

    [Fact]
    public void Hyperlink_ExternalUrl_RoundTrips()
    {
        // The URL is angle-bracketed on the way back out (bug 3f) -- valid
        // CommonMark, and protects against a space/paren in the URL
        // breaking the link syntax.
        var reverse = ToMarkdown("[link text](https://example.com/)");

        Assert.Contains("[link text](<https://example.com/>)", reverse.Markdown);
    }

    [Fact]
    public void Hyperlink_LabelContainingBrackets_DoesNotBreakSurroundingLinkSyntax()
    {
        // The bug this guards against: a hyperlink whose own visible text
        // contains a literal "[" or "]" (e.g. "note [extra] here") would,
        // without label-specific escaping, render as
        // "[note [extra] here](url)" -- most Markdown parsers misread the
        // inner "]" as closing the outer link early, splitting one link
        // into a broken pair. escapeAllBrackets in EscapeInlineText (set by
        // BuildHyperlink) neutralizes every bracket in the label so this
        // can't happen, regardless of where the brackets fall.
        var reverse = ToMarkdown("[note \\[extra\\] here](https://example.com/)");

        Assert.Contains("[note \\[extra\\] here](<https://example.com/>)", reverse.Markdown);
    }

    [Fact]
    public void Hyperlink_InternalBookmarkAnchor_DegradesToPlainTextWithWarning()
    {
        // A w:hyperlink with an internal w:anchor (bookmark link) instead of
        // an r:id -- no relationship to resolve, must degrade, not throw.
        var body = new Body(
            new Paragraph(
                new Hyperlink(new Run(new Text("bookmark link"))) { Anchor = "SomeBookmark" }));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("bookmark link", reverse.Markdown);
        Assert.DoesNotContain("[bookmark link]", reverse.Markdown);
        Assert.NotEmpty(reverse.Warnings);
    }

    // ---- Escaping ----

    [Fact]
    public void SpecialCharacters_SurviveOoxmlToMarkdownToOoxmlRoundTrip()
    {
        const string original = "Текст із * та _ та | та # символами.";
        var converter = new MarkdownConverter(null);

        var forward = converter.ToOoxml(original);
        var reverse = converter.ToMarkdown(forward.FlatOpc);
        var reforward = converter.ToOoxml(reverse.Markdown);

        using var stream = new MemoryStream(reforward.DocxBytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var bodyText = doc.MainDocumentPart.Document.Body.InnerText;

        Assert.Contains(original, bodyText);
    }

    [Fact]
    public void NormalParagraph_LeadingDash_IsEscapedSoItIsNotReadAsAListMarker()
    {
        var body = new Body(
            new Paragraph(new Run(new Text("- not a list item"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("\\- not a list item", reverse.Markdown);
    }

    [Fact]
    public void NormalParagraph_LeadingOrdinalDigit_IsEscapedSoItIsNotReadAsAnOrderedMarker()
    {
        var body = new Body(
            new Paragraph(new Run(new Text("1. not a list item"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("1\\. not a list item", reverse.Markdown);
    }

    [Fact]
    public void PlainText_AngleBracketCharacters_AreBackslashEscaped()
    {
        // Bug 3a: "<" / ">" are now always-escaped, same as "*"/"_"/"#" --
        // otherwise a raw "<script>" in plain text would round-trip into
        // real Markdown/HTML tag syntax. "[" / "]" are NOT always escaped
        // here -- see PlainText_SquareBrackets_AreOnlyEscapedWhenTheyWouldFormALink
        // and MarkdownEscaper.EscapeInlineText's doc comment for why.
        var body = new Body(
            new Paragraph(new Run(new Text("a < b > c"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("a \\< b \\> c", reverse.Markdown);
    }

    [Fact]
    public void PlainText_SquareBrackets_AreOnlyEscapedWhenTheyWouldFormALink()
    {
        // Bug 3a follow-up: a bare "[1]" citation must round-trip unchanged
        // (see LiveBugRegressionTests.Live2 for why: this app reads
        // "\[...\]" as LaTeX math). But literal text shaped like a real
        // link/image -- "]" immediately followed by "(" or "[" -- must still
        // be neutralized, or it would round-trip into an actual clickable
        // link the author never intended to create.
        var body = new Body(
            new Paragraph(new Run(new Text("citation [1] here"))),
            new Paragraph(new Run(new Text("[click](javascript:alert(1))"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("citation [1] here", reverse.Markdown);
        Assert.Contains("[click\\](javascript:alert(1))", reverse.Markdown);
    }

    [Fact]
    public void TextAfterHardBreak_LeadingDash_IsEscaped()
    {
        // Bug 3c: a hard break (Shift+Enter) starts a new physical Markdown
        // line -- text right after one needs the same leading-marker escape
        // as the start of the paragraph, or it can be misread as a list
        // marker on its own line.
        var body = new Body(
            new Paragraph(
                new Run(new Text("line one")),
                new Run(new Break()),
                new Run(new Text("- looks like a list"))));

        var flatOpc = BuildFlatOpc(body);
        var reverse = new MarkdownConverter(null).ToMarkdown(flatOpc);

        Assert.Contains("line one\\\n\\- looks like a list", reverse.Markdown);
    }

    // ---- Thematic break ----

    [Fact]
    public void ThematicBreak_RoundTrips()
    {
        var reverse = ToMarkdown("before\n\n---\n\nafter");

        Assert.Contains("---", reverse.Markdown);
        Assert.Contains("before", reverse.Markdown);
        Assert.Contains("after", reverse.Markdown);
    }
}
