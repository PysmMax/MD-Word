using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MdWord.Core.Math;
using MathOfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Maps a paragraph's run-level OOXML content back to a Markdown inline
/// string. Reverse-direction mirror of <c>MdToOoxml.InlineRunBuilder</c>.
///
/// Adjacent runs with identical bold/italic/strike/code formatting are
/// aggregated into a single delimited span before emitting (per the Phase 4
/// brief) rather than emitting one delimiter pair per run — otherwise two
/// consecutive bold runs that differ only in some formatting-irrelevant
/// property would round-trip as <c>**a****b**</c> instead of <c>**ab**</c>.
/// </summary>
internal static class InlineMarkdownBuilder
{
    private static readonly HashSet<string> MonospaceFonts =
        new(StringComparer.OrdinalIgnoreCase) { "Consolas", "Courier New", "Courier" };

    private readonly struct Format : IEquatable<Format>
    {
        public readonly bool Bold;
        public readonly bool Italic;
        public readonly bool Strike;

        public Format(bool bold, bool italic, bool strike)
        {
            Bold = bold;
            Italic = italic;
            Strike = strike;
        }

        public bool Equals(Format other) => Bold == other.Bold && Italic == other.Italic && Strike == other.Strike;
        public override bool Equals(object obj) => obj is Format other && Equals(other);
        public override int GetHashCode() => (Bold, Italic, Strike).GetHashCode();
    }

    private enum AtomKind
    {
        Text,
        Break,
        Opaque,
    }

    private readonly struct Atom
    {
        public readonly AtomKind Kind;
        public readonly string Text;
        public readonly Format Format;
        public readonly bool IsCode;

        private Atom(AtomKind kind, string text, Format format, bool isCode)
        {
            Kind = kind;
            Text = text;
            Format = format;
            IsCode = isCode;
        }

        public static Atom TextAtom(string text, Format format, bool isCode) => new(AtomKind.Text, text, format, isCode);
        public static Atom BreakAtom() => new(AtomKind.Break, null, default, false);
        public static Atom OpaqueAtom(string rendered) => new(AtomKind.Opaque, rendered, default, false);
    }

    /// <summary>Builds the full Markdown inline text for one paragraph's content.</summary>
    public static string BuildParagraphText(Paragraph paragraph, MdConversionContext context)
    {
        var atoms = new List<Atom>();
        CollectAtoms(paragraph.ChildElements, context, atoms);
        return Render(atoms);
    }

    /// <summary>
    /// Concatenated raw (unescaped, unformatted) text of a paragraph's runs,
    /// hard breaks rendered as <c>\n</c> — for fenced code block lines, where
    /// no Markdown escaping or emphasis should ever apply.
    /// </summary>
    public static string GetRawText(Paragraph paragraph)
    {
        var builder = new StringBuilder();
        foreach (var run in paragraph.Elements<Run>())
        {
            foreach (var child in run.ChildElements)
            {
                switch (child)
                {
                    case Text text:
                        builder.Append(text.Text);
                        break;
                    case Break:
                        builder.Append('\n');
                        break;
                    case TabChar:
                        builder.Append('\t');
                        break;
                }
            }
        }

        return builder.ToString();
    }

    /// <summary>
    /// True when every run in the paragraph is a "code" run (styleId
    /// CodeInline or a monospace font) and there is at least some real text
    /// — the "foreign document, plain Consolas runs, no CodeBlock style"
    /// case the Phase 4 brief explicitly calls out. Links/inline math in the
    /// same paragraph disqualify it (ambiguous, not worth guessing at).
    /// </summary>
    public static bool IsWholeParagraphCode(Paragraph paragraph)
    {
        if (paragraph.Elements<Hyperlink>().Any() || paragraph.Elements<MathOfficeMath>().Any())
        {
            return false;
        }

        var runs = paragraph.Elements<Run>().ToList();
        if (runs.Count == 0)
        {
            return false;
        }

        var hasText = false;
        foreach (var run in runs)
        {
            if (!IsCodeRun(run.RunProperties))
            {
                return false;
            }

            if (run.Elements<Text>().Any(t => !string.IsNullOrEmpty(t.Text)))
            {
                hasText = true;
            }
        }

        return hasText;
    }

    private static void CollectAtoms(IEnumerable<OpenXmlElement> children, MdConversionContext context, List<Atom> atoms)
    {
        foreach (var child in children)
        {
            switch (child)
            {
                case Run run:
                    CollectRunAtoms(run, context, atoms);
                    break;
                case Hyperlink hyperlink:
                    atoms.Add(Atom.OpaqueAtom(BuildHyperlink(hyperlink, context)));
                    break;
                case MathOfficeMath officeMath:
                    atoms.Add(Atom.OpaqueAtom(BuildInlineMath(officeMath, context)));
                    break;
                case InsertedRun insertedRun:
                    // Tracked-change insertion: its content is current
                    // document content, just wrapped — read through it.
                    CollectAtoms(insertedRun.ChildElements, context, atoms);
                    break;
                case DeletedRun:
                    // Tracked-change deletion: not part of the current
                    // document state — skip silently (also its Run children
                    // carry DeletedText, not Text, so reading through would
                    // need separate handling anyway).
                    break;
                case Drawing:
                    context.Warnings.Add("Image skipped (not supported).");
                    break;
                case SdtRun:
                    // Inline content control -- skip per the initial plan's explicit
                    // skip-with-warning list, same as the block-level SdtBlock
                    // case (DocumentMarkdownBuilder's default warning branch).
                    context.Warnings.Add("Content control skipped (not supported).");
                    break;

                // Bookmarks, proofErr, smart tags, rsid-only noise, etc. —
                // nothing this walker needs to read; silently skipped.
            }
        }
    }

    private static void CollectRunAtoms(Run run, MdConversionContext context, List<Atom> atoms)
    {
        var isCode = IsCodeRun(run.RunProperties);
        var format = ReadFormat(run.RunProperties);

        foreach (var child in run.ChildElements)
        {
            switch (child)
            {
                case Text text:
                    atoms.Add(Atom.TextAtom(text.Text, format, isCode));
                    break;
                case FootnoteReference:
                    context.Warnings.Add("Footnote skipped (not supported).");
                    break;
                case EndnoteReference:
                    context.Warnings.Add("Endnote skipped (not supported).");
                    break;
                case Break br:
                    // Only a text-wrapping break (Shift+Enter) maps to a Markdown
                    // line break. A page/column break (Ctrl+Enter) has no Markdown
                    // equivalent — emitting one here left a stray "\" at page
                    // boundaries (LIVE-3). Drop it: content continues seamlessly.
                    if (br.Type is null || br.Type.Value == BreakValues.TextWrapping)
                    {
                        atoms.Add(Atom.BreakAtom());
                    }
                    break;
                case TabChar:
                    atoms.Add(Atom.TextAtom("\t", format, isCode));
                    break;
                case SymbolChar sym:
                    var resolved = SymbolFontMap.Resolve(sym, context);
                    if (resolved.Length > 0)
                    {
                        atoms.Add(Atom.TextAtom(resolved, format, isCode));
                    }
                    break;
            }
        }
    }

    private static string Render(List<Atom> atoms)
    {
        var builder = new StringBuilder();
        var i = 0;

        while (i < atoms.Count)
        {
            var atom = atoms[i];

            if (atom.Kind == AtomKind.Break)
            {
                builder.Append("\\\n");
                i++;
                continue;
            }

            if (atom.Kind == AtomKind.Opaque)
            {
                builder.Append(atom.Text);
                i++;
                continue;
            }

            var format = atom.Format;
            var isCode = atom.IsCode;
            var textBuilder = new StringBuilder(atom.Text);
            i++;

            while (i < atoms.Count && atoms[i].Kind == AtomKind.Text && atoms[i].Format.Equals(format) && atoms[i].IsCode == isCode)
            {
                textBuilder.Append(atoms[i].Text);
                i++;
            }

            builder.Append(RenderTextGroup(textBuilder.ToString(), format, isCode));
        }

        return builder.ToString();
    }

    private static string RenderTextGroup(string text, Format format, bool isCode)
    {
        if (isCode)
        {
            return WrapCode(text);
        }

        var escaped = MarkdownEscaper.EscapeInlineText(text);
        return ApplyFormat(escaped, format);
    }

    private static string WrapCode(string text) =>
        text.Contains('`') ? "`` " + text + " ``" : "`" + text + "`";

    private static string ApplyFormat(string text, Format format)
    {
        if (!format.Bold && !format.Italic && !format.Strike)
        {
            return text;
        }

        // CommonMark emphasis delimiters must hug non-space characters:
        // "**word **" (space before the closing **) is NOT bold. Move any
        // leading/trailing whitespace outside the markers.
        var trimmed = text.Trim();
        if (trimmed.Length == 0)
        {
            return text; // whitespace-only run: nothing to emphasize
        }

        var leading = text.Substring(0, text.Length - text.TrimStart().Length);
        var trailing = text.Substring(text.TrimEnd().Length);

        var result = trimmed;
        if (format.Bold && format.Italic)
        {
            result = "***" + result + "***";
        }
        else if (format.Bold)
        {
            result = "**" + result + "**";
        }
        else if (format.Italic)
        {
            result = "*" + result + "*";
        }

        if (format.Strike)
        {
            result = "~~" + result + "~~";
        }

        return leading + result + trailing;
    }

    private static string BuildHyperlink(Hyperlink hyperlink, MdConversionContext context)
    {
        var innerAtoms = new List<Atom>();
        CollectAtoms(hyperlink.ChildElements, context, innerAtoms);
        var label = Render(innerAtoms);

        var relationshipId = hyperlink.Id?.Value;
        if (relationshipId != null && context.MainPart != null)
        {
            var relationship = context.MainPart.HyperlinkRelationships
                .FirstOrDefault(r => r.Id == relationshipId);

            if (relationship != null)
            {
                return $"[{label}]({relationship.Uri})";
            }
        }

        context.Warnings.Add("An internal link (bookmark) with no external address — inserted as plain text.");
        return label;
    }

    private static string BuildInlineMath(MathOfficeMath officeMath, MdConversionContext context)
    {
        string latex = null;
        string failureReason = null;

        if (context.MathActive && OmmlToLatexConverter.TryConvert(officeMath.OuterXml, context.XslPaths, out latex, out failureReason))
        {
            return "$" + latex + "$";
        }

        context.Warnings.Add(
            $"Could not convert the formula to LaTeX ({failureReason ?? "OMML2MML.XSL unavailable"}) — inserted as `[formula]`.");
        return "`[formula]`";
    }

    private static bool IsCodeRun(RunProperties runProperties)
    {
        if (runProperties == null)
        {
            return false;
        }

        if (string.Equals(runProperties.RunStyle?.Val?.Value, "CodeInline", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var font = runProperties.RunFonts?.Ascii?.Value ?? runProperties.RunFonts?.HighAnsi?.Value;
        return font != null && MonospaceFonts.Contains(font);
    }

    private static Format ReadFormat(RunProperties runProperties)
    {
        if (runProperties == null)
        {
            return default;
        }

        return new Format(
            runProperties.Bold != null && runProperties.Bold.Val?.Value != false,
            runProperties.Italic != null && runProperties.Italic.Val?.Value != false,
            runProperties.Strike != null && runProperties.Strike.Val?.Value != false);
    }
}
