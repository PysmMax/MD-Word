using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MdWord.Core.Math;
using MathOfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Top-level entry point for the reverse (OOXML → Markdown) direction —
/// walks <see cref="Body"/>'s children in order and emits Markdown, mirror
/// of <c>MdToOoxml.WordPackageBuilder</c>/<c>BlockWalker</c> combined. Word
/// stores lists as a flat sequence of paragraphs (not a nested block
/// structure like Markdig's AST), so this walker buffers contiguous runs of
/// list/quote/code paragraphs and flushes them as a single Markdown block
/// when a different kind of paragraph (or a table) interrupts the run.
/// </summary>
internal static class DocumentMarkdownBuilder
{
    public static (string Markdown, string[] Warnings) Build(WordprocessingDocument document, MathXslPaths xslPaths)
    {
        var warnings = new List<string>();
        var mainPart = document.MainDocumentPart;
        var body = mainPart?.Document?.Body;

        var styleCatalog = new StyleCatalog(mainPart?.StyleDefinitionsPart?.Styles);
        var numberingCatalog = new NumberingCatalog(mainPart?.NumberingDefinitionsPart?.Numbering);
        var context = new MdConversionContext(mainPart, xslPaths, styleCatalog, numberingCatalog, warnings);

        if (body == null)
        {
            return (string.Empty, warnings.ToArray());
        }

        var markdown = new Walker(context).WalkBody(body);
        return (markdown, warnings.ToArray());
    }

    private sealed class Walker
    {
        private readonly MdConversionContext _context;
        private readonly List<string> _blocks = new();
        private readonly List<string> _listLines = new();
        private readonly List<string> _quoteLines = new();
        private readonly List<string> _codeLines = new();
        private readonly OrderedListCounters _orderedCounters = new();

        public Walker(MdConversionContext context)
        {
            _context = context;
        }

        public string WalkBody(Body body)
        {
            foreach (var element in body.ChildElements)
            {
                switch (element)
                {
                    case Paragraph paragraph:
                        ProcessParagraph(paragraph);
                        break;
                    case Table table:
                        FlushAll();
                        var tableMarkdown = TableMarkdownBuilder.Build(table, _context);
                        if (tableMarkdown != null)
                        {
                            _blocks.Add(tableMarkdown);
                        }

                        break;
                    case SectionProperties:
                        // The forward direction deliberately never emits one
                        // (so pasting doesn't drag in a section break) — a
                        // foreign document's own trailing sectPr carries no
                        // Markdown-relevant content either way.
                        break;
                    case BookmarkStart:
                    case BookmarkEnd:
                        // Bookmarks carry no Markdown-visible content and Word
                        // plants hidden ones (_GoBack) in most documents --
                        // warning on them would just train users to ignore
                        // real warnings. Same silent-skip the run-level walker
                        // already applies.
                        break;
                    default:
                        _context.Warnings.Add($"Елемент '{element.LocalName}' пропущено (не підтримується).");
                        break;
                }
            }

            FlushAll();

            var joined = string.Join("\n\n", _blocks).TrimEnd();
            return joined.Length == 0 ? string.Empty : joined + "\n";
        }

        private void ProcessParagraph(Paragraph paragraph)
        {
            if (TryGetMathBlockOfficeMaths(paragraph, out var mathBlockOfficeMaths))
            {
                FlushAll();
                foreach (var officeMath in mathBlockOfficeMaths)
                {
                    _blocks.Add(BuildMathBlock(officeMath));
                }

                return;
            }

            var headingLevel = HeadingResolver.ResolveLevel(paragraph, _context.Styles);
            if (headingLevel.HasValue)
            {
                FlushAll();
                var headingText = InlineMarkdownBuilder.BuildParagraphText(paragraph, _context);
                _blocks.Add(new string('#', headingLevel.Value) + " " + headingText);
                return;
            }

            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

            if (string.Equals(styleId, "CodeBlock", StringComparison.OrdinalIgnoreCase))
            {
                FlushList();
                FlushQuote();
                _codeLines.Add(InlineMarkdownBuilder.GetRawText(paragraph));
                return;
            }

            if (InlineMarkdownBuilder.IsWholeParagraphCode(paragraph))
            {
                FlushList();
                FlushQuote();
                _codeLines.Add(InlineMarkdownBuilder.GetRawText(paragraph));
                return;
            }

            if (string.Equals(styleId, "Quote", StringComparison.OrdinalIgnoreCase))
            {
                FlushList();
                FlushCode();
                var quoteText = InlineMarkdownBuilder.BuildParagraphText(paragraph, _context);
                _quoteLines.Add("> " + quoteText);
                return;
            }

            var numberingProperties = paragraph.ParagraphProperties?.NumberingProperties;
            var numId = numberingProperties?.NumberingId?.Val?.Value ?? 0;

            if (numberingProperties != null && numId > 0)
            {
                FlushQuote();
                FlushCode();

                var ilvl = numberingProperties.NumberingLevelReference?.Val?.Value ?? 0;

                if (!_context.Numbering.TryResolve(numId, ilvl, out var isOrdered))
                {
                    _context.Warnings.Add($"Невідомий numId={numId} — оброблено як маркований список.");
                    isOrdered = false;
                }

                var indent = new string(' ', ilvl * 2);
                var marker = isOrdered ? _orderedCounters.Next(numId, ilvl) + "." : "-";
                var itemText = InlineMarkdownBuilder.BuildParagraphText(paragraph, _context);
                _listLines.Add(indent + marker + " " + itemText);
                return;
            }

            if (IsThematicBreak(paragraph))
            {
                FlushAll();
                _blocks.Add("---");
                return;
            }

            // Normal paragraph.
            FlushAll();
            var plainText = InlineMarkdownBuilder.BuildParagraphText(paragraph, _context);
            if (string.IsNullOrEmpty(plainText))
            {
                return;
            }

            _blocks.Add(MarkdownEscaper.EscapeLeadingMarker(plainText));
        }

        private void FlushAll()
        {
            FlushList();
            FlushQuote();
            FlushCode();
        }

        private void FlushList()
        {
            if (_listLines.Count == 0)
            {
                return;
            }

            _blocks.Add(string.Join("\n", _listLines));
            _listLines.Clear();
            _orderedCounters.Reset();
        }

        private void FlushQuote()
        {
            if (_quoteLines.Count == 0)
            {
                return;
            }

            _blocks.Add(string.Join("\n>\n", _quoteLines));
            _quoteLines.Clear();
        }

        private void FlushCode()
        {
            if (_codeLines.Count == 0)
            {
                return;
            }

            _blocks.Add("```\n" + string.Join("\n", _codeLines) + "\n```");
            _codeLines.Clear();
        }

        private string BuildMathBlock(MathOfficeMath officeMath)
        {
            string latex = null;
            string failureReason = null;

            if (_context.MathActive && OmmlToLatexConverter.TryConvert(officeMath.OuterXml, _context.XslPaths, out latex, out failureReason))
            {
                return "$$\n" + latex + "\n$$";
            }

            _context.Warnings.Add(
                $"Формулу (блок) не вдалося конвертувати в LaTeX ({failureReason ?? "OMML2MML.XSL недоступний"}) — вставлено як `[formula]`.");
            return "`[formula]`";
        }

        private static bool TryGetMathBlockOfficeMaths(Paragraph paragraph, out List<MathOfficeMath> officeMaths)
        {
            officeMaths = null;
            var mathParagraph = paragraph.Elements<MathParagraph>().FirstOrDefault();
            if (mathParagraph == null)
            {
                return false;
            }

            var maths = mathParagraph.Elements<MathOfficeMath>().ToList();
            if (maths.Count == 0)
            {
                return false;
            }

            officeMaths = maths;
            return true;
        }

        private static bool IsThematicBreak(Paragraph paragraph)
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties?.ParagraphBorders?.BottomBorder == null)
            {
                return false;
            }

            if (paragraphProperties.ParagraphStyleId != null || paragraphProperties.NumberingProperties != null)
            {
                return false;
            }

            return !paragraph.Elements<Run>().Any(run => run.Elements<Text>().Any(text => !string.IsNullOrEmpty(text.Text)));
        }
    }
}
