using DocumentFormat.OpenXml.Wordprocessing;

namespace MdWord.Core.MdToOoxml;

/// <summary>
/// Builds the numbering.xml part backing bullet and ordered lists: two
/// <c>abstractNum</c> definitions (bullet, ordered), each with 3 nesting
/// levels, and two <c>num</c> instances referencing them. List paragraphs
/// reference <see cref="BulletNumId"/>/<see cref="OrderedNumId"/> via
/// <c>numPr</c>/<c>ilvl</c> (see <see cref="BlockWalker"/>).
///
/// Known Phase 1 limitation (accepted per PLAN.md §Phase 1, which specifies
/// exactly two abstractNum definitions, not one per list instance): every
/// ordered list in a document shares the single <see cref="OrderedNumId"/>
/// <c>num</c> instance, so if a document has two *separate* top-level
/// ordered lists, Word continues numbering across them (e.g. 1,2,3 then
/// 4,5,6) instead of each restarting at 1. Locked in by
/// MdWord.Core.Tests.BlockMappingTests.
/// TwoSeparateOrderedLists_ShareOneNumId_KnownPhase1NumberingLimitation.
/// Revisit only if this proves to matter in practice (would need a fresh
/// <c>num</c> instance per top-level ListBlock instead of one static pair).
/// </summary>
internal static class NumberingPartBuilder
{
    public const int BulletAbstractNumId = 0;
    public const int OrderedAbstractNumId = 1;
    public const int BulletNumId = 1;
    public const int OrderedNumId = 2;

    /// <summary>Deepest zero-based <c>ilvl</c> the generated numbering supports (3 levels: 0,1,2).</summary>
    public const int MaxLevelIndex = 2;

    public static Numbering BuildMinimalNumbering()
    {
        var numbering = new Numbering();

        numbering.Append(BuildAbstractNum(BulletAbstractNumId, isOrdered: false));
        numbering.Append(BuildAbstractNum(OrderedAbstractNumId, isOrdered: true));

        numbering.Append(new NumberingInstance(new AbstractNumId { Val = BulletAbstractNumId }) { NumberID = BulletNumId });
        numbering.Append(new NumberingInstance(new AbstractNumId { Val = OrderedAbstractNumId }) { NumberID = OrderedNumId });

        return numbering;
    }

    private static AbstractNum BuildAbstractNum(int abstractNumId, bool isOrdered)
    {
        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };

        for (var level = 0; level <= MaxLevelIndex; level++)
        {
            abstractNum.Append(isOrdered ? BuildOrderedLevel(level) : BuildBulletLevel(level));
        }

        return abstractNum;
    }

    private static Level BuildBulletLevel(int levelIndex)
    {
        var bulletChar = levelIndex switch
        {
            0 => "•", // •
            1 => "◦", // ◦
            _ => "▪", // ▪
        };

        return new Level(
            new NumberingFormat { Val = NumberFormatValues.Bullet },
            new LevelText { Val = bulletChar },
            new LevelJustification { Val = LevelJustificationValues.Left },
            new PreviousParagraphProperties(
                new Indentation { Left = ((levelIndex + 1) * 720).ToString(), Hanging = "360" }))
        {
            LevelIndex = levelIndex,
        };
    }

    private static Level BuildOrderedLevel(int levelIndex)
    {
        return new Level(
            new StartNumberingValue { Val = 1 },
            new NumberingFormat { Val = NumberFormatValues.Decimal },
            new LevelText { Val = $"%{levelIndex + 1}." },
            new LevelJustification { Val = LevelJustificationValues.Left },
            new PreviousParagraphProperties(
                new Indentation { Left = ((levelIndex + 1) * 720).ToString(), Hanging = "360" }))
        {
            LevelIndex = levelIndex,
        };
    }
}
