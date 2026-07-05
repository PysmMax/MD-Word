using System.Text;
using System.Text.RegularExpressions;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Escapes Markdown-significant characters in literal run text so a
/// round-trip through a Markdown renderer doesn't reinterpret plain prose as
/// structure (Phase 4 brief). Two separate concerns:
/// <list type="bullet">
/// <item><see cref="EscapeInlineText"/> — characters that are significant
/// <i>anywhere</i> they appear (<c>* _ # | \ `</c>).</item>
/// <item><see cref="EscapeLeadingMarker"/> — characters that are only
/// significant as the very first thing on a line (<c>-</c>/<c>N.</c> as a
/// list marker, <c>&gt;</c> as a blockquote marker) — applied once to a
/// whole assembled paragraph's text, not per character, since a mid-sentence
/// "5 &gt; 3" or "open-source" must not be escaped.</item>
/// </list>
/// </summary>
internal static class MarkdownEscaper
{
    private static readonly Regex LeadingDash = new(@"^-(\s|$)", RegexOptions.Compiled);
    private static readonly Regex LeadingOrdered = new(@"^(\d+)\.(\s|$)", RegexOptions.Compiled);

    /// <summary>Backslash-escapes <c>* _ # | \ `</c> wherever they occur.</summary>
    public static string EscapeInlineText(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        var builder = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '\\':
                case '*':
                case '_':
                case '#':
                case '|':
                case '`':
                    builder.Append('\\');
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
    /// Escapes a leading <c>&gt;</c>, <c>- </c>, or <c>N. </c> at the start of
    /// an already-<see cref="EscapeInlineText"/>-escaped paragraph string, so
    /// it isn't misread as a blockquote/list marker. Must only be called on
    /// plain (non-list, non-quote) paragraph text — list items and
    /// blockquotes are expected to start with those markers.
    /// </summary>
    public static string EscapeLeadingMarker(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        if (text[0] == '>')
        {
            return "\\" + text;
        }

        if (LeadingDash.IsMatch(text))
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
}
