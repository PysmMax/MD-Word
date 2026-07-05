using System.Text;
using System.Text.RegularExpressions;

namespace MdWord.Core.Math;

/// <summary>
/// Runs on the raw Markdown string before <c>Markdig.Parse</c>: rewrites
/// AI-tool-style LaTeX delimiters (<c>\(...\)</c>, <c>\[...\]</c>) to the
/// <c>$...$</c>/<c>$$...$$</c> form Markdig's mathematics extension actually
/// recognizes. Handles two shapes: (a) same-line pairs, substituted within
/// each line independently; (b) a lone <c>\[</c> line paired with a later
/// lone <c>\]</c> line (the usual multi-line display-math emitted by AI
/// tools), rewritten to <c>$$</c> fence lines with the body kept verbatim.
/// Tracks fenced-code-block state (``` / ~~~) so code samples are never
/// touched; an unpaired lone <c>\[</c> is left as-is.
/// </summary>
internal static class LatexDelimiterPreprocessor
{
    private static readonly Regex FenceMarker = new(@"^\s{0,3}(`{3,}|~{3,})", RegexOptions.Compiled);
    private static readonly Regex ParenDelimiter = new(@"\\\((.*?)\\\)", RegexOptions.Compiled);
    private static readonly Regex BracketDelimiter = new(@"\\\[(.*?)\\\]", RegexOptions.Compiled);
    private static readonly Regex LoneBracketOpen = new(@"^\s{0,3}\\\[\s*$", RegexOptions.Compiled);
    private static readonly Regex LoneBracketClose = new(@"^\s{0,3}\\\]\s*$", RegexOptions.Compiled);

    public static string Rewrite(string markdown)
    {
        var lines = markdown.Split('\n');

        // Pass 1: fenced-code state per line (``` / ~~~), marker lines included.
        var inFence = new bool[lines.Length];
        string openFence = null;
        for (var i = 0; i < lines.Length; i++)
        {
            var fenceMatch = FenceMarker.Match(lines[i]);
            if (fenceMatch.Success)
            {
                var marker = fenceMatch.Groups[1].Value;
                if (openFence == null)
                {
                    openFence = marker;
                }
                else if (marker[0] == openFence[0] && marker.Length >= openFence.Length)
                {
                    openFence = null;
                }

                inFence[i] = true; // fence marker lines are never rewritten
            }
            else
            {
                inFence[i] = openFence != null;
            }
        }

        // Pass 2: pair lone "\[" / "\]" lines outside fences.
        var isDisplayDelimiter = new bool[lines.Length];
        var inDisplayMath = new bool[lines.Length];
        var openIndex = -1;
        for (var i = 0; i < lines.Length; i++)
        {
            if (inFence[i])
            {
                continue;
            }

            if (openIndex < 0 && LoneBracketOpen.IsMatch(lines[i]))
            {
                openIndex = i;
            }
            else if (openIndex >= 0 && LoneBracketClose.IsMatch(lines[i]))
            {
                isDisplayDelimiter[openIndex] = true;
                isDisplayDelimiter[i] = true;
                for (var j = openIndex + 1; j < i; j++)
                {
                    inDisplayMath[j] = true;
                }

                openIndex = -1;
            }
        }

        // Pass 3: emit.
        var result = new StringBuilder(markdown.Length);
        for (var i = 0; i < lines.Length; i++)
        {
            if (isDisplayDelimiter[i])
            {
                result.Append("$$");
            }
            else if (inFence[i] || inDisplayMath[i])
            {
                result.Append(lines[i]);
            }
            else
            {
                result.Append(RewriteLine(lines[i]));
            }

            if (i < lines.Length - 1)
            {
                result.Append('\n');
            }
        }

        return result.ToString();
    }

    private static string RewriteLine(string line)
    {
        // Regex replacement-string escaping: "$$" emits one literal '$';
        // "$1" substitutes capture group 1. So "$$" + "$1" + "$$" (written
        // below without the internal concatenation) yields "$<content>$",
        // and doubled up yields "$$<content>$$".
        line = ParenDelimiter.Replace(line, "$$$1$$");
        line = BracketDelimiter.Replace(line, "$$$$$1$$$$");
        return line;
    }
}
