using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;
using MdWord.Core.Math;
using MdTable = Markdig.Extensions.Tables.Table;
using MdTableCell = Markdig.Extensions.Tables.TableCell;
using MdTableRow = Markdig.Extensions.Tables.TableRow;
using MathOfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;

namespace MdWord.Core.MdToOoxml;

/// <summary>
/// Maps top-level Markdig blocks to OOXML body content. Grows one block type
/// per PLAN.md Phase 1 increment; see <see cref="InlineRunBuilder"/> for the
/// shared inline (run-level) mapping every block delegates to.
/// </summary>
internal static class BlockWalker
{
    public static IEnumerable<OpenXmlElement> MapBlocks(MarkdownDocument markdownDocument, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        foreach (var block in markdownDocument)
        {
            switch (block)
            {
                case HeadingBlock { Inline: { } headingInline } heading:
                    yield return BuildStyledParagraph($"Heading{heading.Level}", headingInline, mainPart, mathContext);
                    break;
                case ParagraphBlock { Inline: { } paragraphInline }:
                    yield return BuildStyledParagraph("Normal", paragraphInline, mainPart, mathContext);
                    break;
                case ThematicBreakBlock:
                    yield return BuildThematicBreakParagraph();
                    break;
                case QuoteBlock quote:
                    foreach (var quoteParagraph in MapQuoteParagraphs(quote, mainPart, mathContext))
                    {
                        yield return quoteParagraph;
                    }

                    break;
                case Markdig.Extensions.Mathematics.MathBlock mathBlock:
                    // MathBlock derives from CodeBlock in Markdig, so this case must
                    // stay ahead of the plain CodeBlock one below. Converts to OMML
                    // (m:oMathPara) when mathContext.IsActive; otherwise degrades to
                    // the literal "$$...$$" source so content is never silently lost.
                    yield return BuildMathBlockParagraph(mathBlock, mathContext);
                    break;
                case CodeBlock codeBlock:
                    foreach (var codeParagraph in MapCodeBlockParagraphs(codeBlock))
                    {
                        yield return codeParagraph;
                    }

                    break;
                case ListBlock list:
                    foreach (var listParagraph in MapListParagraphs(list, level: 0, mainPart, mathContext))
                    {
                        yield return listParagraph;
                    }

                    break;
                case MdTable table:
                    yield return BuildTable(table, mainPart, mathContext);
                    break;
                case HtmlBlock htmlBlock:
                    // Per PLAN.md Phase 1: "HtmlInline/HtmlBlock → literal-текст
                    // (безпечна деградація)" — raw HTML is never dropped, just shown
                    // as its own source text.
                    yield return BuildHtmlBlockParagraph(htmlBlock);
                    break;
                case Markdig.Syntax.LinkReferenceDefinitionGroup:
                    // Reference definitions ([id]: url) legitimately produce no
                    // visible output -- silent skip is correct, not content loss.
                    break;
                default:
                    // Never drop content silently (PLAN.md §6) -- at minimum warn.
                    mathContext.Warnings.Add($"Блок '{block.GetType().Name}' пропущено (не підтримується).");
                    break;
            }
        }
    }

    private static Paragraph BuildStyledParagraph(string styleId, Markdig.Syntax.Inlines.ContainerInline inline, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = styleId }));
        paragraph.Append(InlineRunBuilder.Build(inline, mainPart, mathContext));
        return paragraph;
    }

    private static IEnumerable<OpenXmlElement> MapQuoteParagraphs(QuoteBlock quote, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        foreach (var nested in quote)
        {
            switch (nested)
            {
                case ParagraphBlock { Inline: { } inline }:
                    yield return BuildStyledParagraph("Quote", inline, mainPart, mathContext);
                    break;
                case ListBlock nestedList:
                    foreach (var element in MapListParagraphs(nestedList, level: 0, mainPart, mathContext))
                    {
                        yield return element;
                    }

                    break;
                case Markdig.Extensions.Mathematics.MathBlock mathBlock:
                    // MathBlock derives from CodeBlock -- must stay ahead of it.
                    yield return BuildMathBlockParagraph(mathBlock, mathContext);
                    break;
                case CodeBlock codeBlock:
                    foreach (var codeParagraph in MapCodeBlockParagraphs(codeBlock))
                    {
                        yield return codeParagraph;
                    }

                    break;
                case QuoteBlock nestedQuote:
                    foreach (var element in MapQuoteParagraphs(nestedQuote, mainPart, mathContext))
                    {
                        yield return element;
                    }

                    break;
                case MdTable table:
                    yield return BuildTable(table, mainPart, mathContext);
                    break;
                default:
                    // Never drop content silently (PLAN.md §6) -- at minimum warn.
                    mathContext.Warnings.Add($"Блок '{nested.GetType().Name}' всередині цитати пропущено (не підтримується).");
                    break;
            }
        }
    }

    private static IEnumerable<Paragraph> MapCodeBlockParagraphs(CodeBlock codeBlock)
    {
        for (var i = 0; i < codeBlock.Lines.Count; i++)
        {
            var text = codeBlock.Lines.Lines[i].Slice.ToString();
            var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "CodeBlock" }));
            paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            yield return paragraph;
        }
    }

    private static Paragraph BuildMathBlockParagraph(Markdig.Extensions.Mathematics.MathBlock mathBlock, MathConversionContext mathContext)
    {
        var tex = ExtractMathBlockTex(mathBlock);

        if (mathContext.IsActive)
        {
            if (MathConverter.TryConvert(tex, displayMode: true, mathContext.XslPaths, out var ommlOuterXml, out var failureReason))
            {
                var officeMath = new MathOfficeMath(ommlOuterXml);
                return new Paragraph(new MathParagraph(officeMath));
            }

            mathContext.Warnings.Add($"Формулу `{tex}` вставлено як текст: {failureReason}");
        }

        return BuildLiteralMathBlockParagraph(tex);
    }

    private static string ExtractMathBlockTex(Markdig.Extensions.Mathematics.MathBlock mathBlock)
    {
        var lines = new List<string>();
        for (var i = 0; i < mathBlock.Lines.Count; i++)
        {
            lines.Add(mathBlock.Lines.Lines[i].Slice.ToString());
        }

        return string.Join("\n", lines);
    }

    private static Paragraph BuildLiteralMathBlockParagraph(string tex)
    {
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Normal" }));
        paragraph.Append(new Run(new Text("$$") { Space = SpaceProcessingModeValues.Preserve }));

        foreach (var line in tex.Split('\n'))
        {
            var run = new Run(new Break(), new Text(line) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);
        }

        paragraph.Append(new Run(new Break(), new Text("$$") { Space = SpaceProcessingModeValues.Preserve }));

        return paragraph;
    }

    private static Paragraph BuildHtmlBlockParagraph(HtmlBlock htmlBlock)
    {
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Normal" }));

        for (var i = 0; i < htmlBlock.Lines.Count; i++)
        {
            var text = htmlBlock.Lines.Lines[i].Slice.ToString();
            var run = i == 0
                ? new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })
                : new Run(new Break(), new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.Append(run);
        }

        return paragraph;
    }

    private static IEnumerable<OpenXmlElement> MapListParagraphs(ListBlock list, int level, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        var numId = list.IsOrdered ? NumberingPartBuilder.OrderedNumId : NumberingPartBuilder.BulletNumId;
        var ilvl = System.Math.Min(level, NumberingPartBuilder.MaxLevelIndex);

        foreach (var item in list)
        {
            if (item is not ListItemBlock listItem)
            {
                continue;
            }

            foreach (var nested in listItem)
            {
                switch (nested)
                {
                    case ParagraphBlock { Inline: { } inline }:
                        yield return BuildListItemParagraph(inline, numId, ilvl, mainPart, mathContext);
                        break;
                    case ListBlock nestedList:
                        foreach (var nestedParagraph in MapListParagraphs(nestedList, level + 1, mainPart, mathContext))
                        {
                            yield return nestedParagraph;
                        }

                        break;
                    case Markdig.Extensions.Mathematics.MathBlock mathBlock:
                        // MathBlock derives from CodeBlock -- must stay ahead of it.
                        yield return BuildMathBlockParagraph(mathBlock, mathContext);
                        break;
                    case CodeBlock codeBlock:
                        foreach (var codeParagraph in MapCodeBlockParagraphs(codeBlock))
                        {
                            yield return codeParagraph;
                        }

                        break;
                    case QuoteBlock nestedQuote:
                        foreach (var element in MapQuoteParagraphs(nestedQuote, mainPart, mathContext))
                        {
                            yield return element;
                        }

                        break;
                    case MdTable table:
                        yield return BuildTable(table, mainPart, mathContext);
                        break;
                    default:
                        // Never drop content silently (PLAN.md §6) -- at minimum warn.
                        mathContext.Warnings.Add($"Блок '{nested.GetType().Name}' всередині пункту списку пропущено (не підтримується).");
                        break;
                }
            }
        }
    }

    private static Paragraph BuildListItemParagraph(Markdig.Syntax.Inlines.ContainerInline inline, int numId, int ilvl, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        var numPr = new NumberingProperties(
            new NumberingLevelReference { Val = ilvl },
            new NumberingId { Val = numId });
        var paragraph = new Paragraph(new ParagraphProperties(numPr));
        paragraph.Append(InlineRunBuilder.Build(inline, mainPart, mathContext));
        return paragraph;
    }

    private static Table BuildTable(MdTable table, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        var wordTable = new Table();

        wordTable.Append(new TableProperties(new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        var columnCount = table.ColumnDefinitions.Count;
        var tableGrid = new TableGrid();
        for (var i = 0; i < columnCount; i++)
        {
            tableGrid.Append(new GridColumn());
        }

        wordTable.Append(tableGrid);

        foreach (var rowBlock in table.OfType<MdTableRow>())
        {
            wordTable.Append(BuildTableRow(rowBlock, mainPart, mathContext));
        }

        return wordTable;
    }

    private static TableRow BuildTableRow(MdTableRow rowBlock, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        var row = new TableRow();

        foreach (var cellBlock in rowBlock.OfType<MdTableCell>())
        {
            row.Append(BuildTableCell(cellBlock, rowBlock.IsHeader, mainPart, mathContext));
        }

        return row;
    }

    private static TableCell BuildTableCell(MdTableCell cellBlock, bool isHeader, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        var cell = new TableCell();

        foreach (var nested in cellBlock)
        {
            if (nested is ParagraphBlock { Inline: { } inline })
            {
                cell.Append(BuildTableCellParagraph(inline, isHeader, mainPart, mathContext));
            }
        }

        if (!cell.Elements<Paragraph>().Any())
        {
            cell.Append(new Paragraph());
        }

        return cell;
    }

    private static Paragraph BuildTableCellParagraph(Markdig.Syntax.Inlines.ContainerInline inline, bool bold, MainDocumentPart mainPart, MathConversionContext mathContext)
    {
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Normal" }));
        var runs = InlineRunBuilder.Build(inline, mainPart, mathContext).ToList();

        if (bold)
        {
            foreach (var run in runs.OfType<Run>())
            {
                run.RunProperties ??= new RunProperties();

                // Typed-setter assignment (not PrependChild) so Bold lands in its
                // schema-mandated position even when the run already carries an
                // earlier-in-sequence child such as rStyle (e.g. a code span).
                run.RunProperties.Bold = new Bold();
            }
        }

        paragraph.Append(runs);
        return paragraph;
    }

    private static Paragraph BuildThematicBreakParagraph()
    {
        var borders = new ParagraphBorders(
            new BottomBorder { Val = BorderValues.Single, Size = 6, Space = 1 });
        return new Paragraph(new ParagraphProperties(borders));
    }
}
