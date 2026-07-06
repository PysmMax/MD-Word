using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Maps a <see cref="Table"/> back to a Markdown pipe table. Reverse mirror
/// of <c>MdToOoxml.BlockWalker.BuildTable</c>. Markdown pipe tables always
/// have exactly one header row (the separator line right after it) — the
/// forward direction only bolds the first row's text, it does not set a
/// dedicated "this is a header row" property, so the first row is always
/// treated as the header on the way back too, matching how
/// <c>samples/demo.md</c> itself is shaped.
///
/// <c>gridSpan</c>/<c>vMerge</c> can't be expressed in a pipe table (per
/// the initial plan) — both degrade to an empty cell in that grid position plus a
/// <see cref="MdConversionContext.Warnings"/> entry, rather than attempting
/// merge-aware reconstruction.
/// </summary>
internal static class TableMarkdownBuilder
{
    public static string Build(Table table, MdConversionContext context)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0)
        {
            return null;
        }

        var columnCount = ComputeColumnCount(table, rows);
        var lines = new List<string>(rows.Count + 1);

        for (var r = 0; r < rows.Count; r++)
        {
            var cells = ExpandRowCells(rows[r], columnCount, context);
            lines.Add(FormatRow(cells));

            if (r == 0)
            {
                lines.Add(FormatRow(Enumerable.Repeat("---", columnCount)));
            }
        }

        return string.Join("\n", lines);
    }

    private static string FormatRow(IEnumerable<string> cells) => "| " + string.Join(" | ", cells) + " |";

    private static int ComputeColumnCount(Table table, List<TableRow> rows)
    {
        var grid = table.GetFirstChild<TableGrid>();
        var gridColumnCount = grid?.Elements<GridColumn>().Count() ?? 0;
        if (gridColumnCount > 0)
        {
            return gridColumnCount;
        }

        return rows.Max(row => row.Elements<TableCell>().Sum(GridSpanOf));
    }

    private static int GridSpanOf(TableCell cell) => cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;

    private static bool IsVerticalMergeContinuation(TableCell cell)
    {
        var verticalMerge = cell.TableCellProperties?.VerticalMerge;
        if (verticalMerge == null)
        {
            return false;
        }

        // <w:vMerge/> with no @w:val means "continue" per the OOXML spec —
        // only an explicit val="restart" starts a new merged region.
        return verticalMerge.Val == null || verticalMerge.Val.Value == MergedCellValues.Continue;
    }

    private static List<string> ExpandRowCells(TableRow row, int columnCount, MdConversionContext context)
    {
        var result = new List<string>();

        foreach (var cell in row.Elements<TableCell>())
        {
            var gridSpan = GridSpanOf(cell);

            if (IsVerticalMergeContinuation(cell))
            {
                context.Warnings.Add("A vertically merged cell (vMerge) was lost — inserted an empty cell.");
                for (var i = 0; i < gridSpan; i++)
                {
                    result.Add(string.Empty);
                }

                continue;
            }

            result.Add(BuildCellContent(cell, context));

            if (gridSpan > 1)
            {
                context.Warnings.Add($"A horizontally merged cell (gridSpan={gridSpan}) was lost — added empty cells.");
                for (var i = 1; i < gridSpan; i++)
                {
                    result.Add(string.Empty);
                }
            }
        }

        while (result.Count < columnCount)
        {
            result.Add(string.Empty);
        }

        if (result.Count > columnCount)
        {
            result = result.Take(columnCount).ToList();
        }

        return result;
    }

    private static string BuildCellContent(TableCell cell, MdConversionContext context)
    {
        var paragraphTexts = cell.Elements<Paragraph>()
            // A hard break renders as "\<newline>" (InlineMarkdownBuilder.Render),
            // which would split this cell's pipe-table row across lines -- inside
            // a cell it must use the same <br> convention as paragraph joins.
            .Select(paragraph => InlineMarkdownBuilder.BuildParagraphText(paragraph, context).Replace("\\\n", "<br>"))
            .Where(text => !string.IsNullOrEmpty(text));

        return string.Join("<br>", paragraphTexts);
    }
}
