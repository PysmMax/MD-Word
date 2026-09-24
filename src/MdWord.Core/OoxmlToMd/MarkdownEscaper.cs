using System;
using System.Text;
using System.Text.RegularExpressions;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Escapes Markdown-significant characters in literal run text so a
/// round-trip through a Markdown renderer doesn't reinterpret plain prose as
/// structure (Phase 4 brief). Two separate concerns:
/// <list type="bullet">
/// <item><see cref="EscapeInlineText"/> — characters that are significant
/// <i>anywhere</i> they appear (<c>* _ # | \ ` &lt; &gt;</c>), plus a
/// targeted escape of <c>]</c> when it is immediately followed by
/// <c>(</c> or <c>[</c> (the only way a stray bracket in plain prose can
/// turn into a real link/image). A bare <c>[</c> or <c>]</c> is left alone:
/// this app treats a backslash-escaped bracket pair (<c>\[...\]</c>) as a
/// LaTeX display-math delimiter on the Markdown-to-Word path
/// (<c>LatexDelimiterPreprocessor</c>), so escaping every bracket here
/// would turn an ordinary citation like "[1]" into math on round-trip —
/// confirmed by <c>LiveBugRegressionTests.Live2</c>.</item>
/// <item><see cref="EscapeLeadingMarker"/> — characters that are only
/// significant as the very first thing on a line (<c>-</c>/<c>+</c>/<c>N.</c>
/// as a list marker, <c>&gt;</c> as a blockquote marker) — applied once to a
/// whole assembled paragraph's text, not per character, since a mid-sentence
/// "5 &gt; 3" or "open-source" must not be escaped.</item>
/// </list>
/// </summary>
internal static class MarkdownEscaper
{
    private static readonly Regex LeadingDash = new(@"^-(\s|$)", RegexOptions.Compiled);
    private static readonly Regex LeadingPlus = new(@"^\+(\s|$)", RegexOptions.Compiled);
    private static readonly Regex LeadingOrdered = new(@"^(\d+)\.(\s|$)", RegexOptions.Compiled);

    /// <summary>
    /// Backslash-escapes <c>* _ # | \ ` ~ ^</c> wherever they occur, and
    /// <c>+</c>/<c>=</c> only when doubled. Since v1.0.2 the parser has all
    /// EmphasisExtras enabled: a PAIR of single <c>~</c>/<c>^</c> anywhere
    /// in a paragraph becomes sub/superscript (verified: "~10 to 20~" parses
    /// as subscript), so both are escaped unconditionally; <c>++</c>/<c>==</c>
    /// only open inserted/marked spans as adjacent pairs, and a lone
    /// <c>+</c>/<c>=</c> is common prose ("a = b + c") — escape each
    /// character that has an identical neighbor, leave singles alone.
    /// </summary>
    /// <param name="text">The raw run text to escape.</param>
    /// <param name="escapeAllBrackets">
    /// When true, every literal <c>[</c> / <c>]</c> is escaped unconditionally
    /// instead of the narrow default rule below. Set only by
    /// <c>InlineMarkdownBuilder.BuildHyperlink</c> for a link's own visible
    /// label text: a stray bracket there can prematurely close (or reopen)
    /// the surrounding <c>[label](url)</c> construct — e.g. a label like
    /// "See [note] here" would otherwise render as
    /// <c>[See [note] here](url)</c>, which most Markdown parsers misread as
    /// two separate links. A hyperlink label is a narrow, rare place for
    /// literal brackets to appear, so the LaTeX-collision risk described
    /// above is accepted there specifically, in exchange for a correctly
    /// formed link. General paragraph prose keeps the narrow default.
    /// </param>
    public static string EscapeInlineText(string text, bool escapeAllBrackets = false)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        var builder = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            switch (ch)
            {
                case '\\':
                case '*':
                case '_':
                case '#':
                case '|':
                case '`':
                case '~':
                case '^':
                case '<':
                case '>':
                    builder.Append('\\');
                    builder.Append(ch);
                    break;
                case '[':
                    if (escapeAllBrackets)
                    {
                        builder.Append('\\');
                    }

                    builder.Append(ch);
                    break;
                case ']':
                    // A bare "]" is harmless; escape it only when it would
                    // actually complete a link/image or reference construct
                    // ("...](" or "...][") right after it, unless the caller
                    // asked for every bracket escaped (see the parameter doc
                    // above). Never escape a bare "[" outside that case.
                    if (escapeAllBrackets
                        || (i + 1 < text.Length && (text[i + 1] == '(' || text[i + 1] == '[')))
                    {
                        builder.Append('\\');
                    }

                    builder.Append(ch);
                    break;
                case '+':
                case '=':
                    var hasSameNeighbor =
                        (i + 1 < text.Length && text[i + 1] == ch)
                        || (i > 0 && text[i - 1] == ch);
                    if (hasSameNeighbor)
                    {
                        builder.Append('\\');
                    }

                    builder.Append(ch);
                    break;
                default:
                    builder.Append(ch);
                    break;
            }
        }

        return builder.ToString();
    }

    /// <summary>
    /// Escapes a leading <c>&gt;</c>, <c>- </c>, <c>+ </c>, or <c>N. </c> at
    /// the start of an already-<see cref="EscapeInlineText"/>-escaped
    /// paragraph string, so it isn't misread as a blockquote/list marker.
    /// Also matters for text that is ITSELF the content of a heading/quote/
    /// list-item paragraph (see <see cref="EscapeLeadingMarkers"/>): e.g. a
    /// blockquote paragraph whose own text starts with "&gt; " would render
    /// as "&gt; &gt; text", which Markdown reads as a nested blockquote, not
    /// literal "&gt;" content.
    /// </summary>
    public static string EscapeLeadingMarker(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        // Now effectively unreachable via the normal call path: ">" is also
        // in EscapeInlineText's always-escape switch above, so by the time
        // an EscapeInlineText'd string reaches here a leading ">" already
        // reads "\>", not ">". Left in place as a harmless safety net (and
        // documentation of the original bug) in case this is ever called on
        // text that hasn't gone through EscapeInlineText first.
        if (text[0] == '>')
        {
            return "\\" + text;
        }

        if (LeadingDash.IsMatch(text) || LeadingPlus.IsMatch(text))
        {
            return "\\" + text;
        }

        var orderedMatch = LeadingOrdered.Match(text);
        if (orderedMatch.Success)
        {
            var digits = orderedMatch.Groups[1].Value;
            return digits + "\\." + text.Substring(digits.Length + 1);
        }

        return text;
    }

    /// <summary>
    /// Applies <see cref="EscapeLeadingMarker"/> to the start of
    /// <paramref name="text"/> and again to the start of every line that
    /// follows a hard line break. <c>InlineMarkdownBuilder.Render</c> emits
    /// a hard break (Shift+Enter) as the literal two-character sequence
    /// <c>"\\\n"</c> (backslash + newline); a hard break starts a new
    /// physical Markdown line, so text right after one is just as exposed to
    /// being misread as a blockquote/list marker as the very first line of
    /// the paragraph. Splitting on that literal sequence is safe: OOXML text
    /// runs never carry a real newline character themselves (Word models
    /// line/paragraph breaks as separate elements, not embedded '\n's), so
    /// every occurrence of "\\\n" in an already-rendered paragraph string is
    /// a hard break, never paragraph prose.
    /// </summary>
    public static string EscapeLeadingMarkers(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        const string hardBreak = "\\\n";
        var lines = text.Split(new[] { hardBreak }, StringSplitOptions.None);
        for (var i = 0; i < lines.Length; i++)
        {
            lines[i] = EscapeLeadingMarker(lines[i]);
        }

        return string.Join(hardBreak, lines);
    }
}
