using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Resolves a paragraph's <c>numId</c>/<c>ilvl</c> (from its
/// <see cref="NumberingProperties"/>) to a bullet-vs-ordered decision by
/// walking the document's own <c>numbering.xml</c> (<see cref="Numbering"/>):
/// <c>numId</c> → <see cref="NumberingInstance"/> → <c>abstractNumId</c> →
/// <see cref="AbstractNum"/> → the <see cref="Level"/> for that <c>ilvl</c> →
/// its <see cref="NumberingFormat"/>. Deliberately does not assume
/// <c>MdToOoxml.NumberingPartBuilder</c>'s own <c>BulletNumId</c>/
/// <c>OrderedNumId</c> constants (1/2) — a foreign document can use any
/// <c>numId</c>, per the Phase 4 brief.
/// </summary>
internal sealed class NumberingCatalog
{
    private readonly Dictionary<int, int> _numIdToAbstractId = new();
    private readonly Dictionary<int, Dictionary<int, NumberFormatValues>> _abstractLevelFormats = new();

    public NumberingCatalog(Numbering numbering)
    {
        if (numbering == null)
        {
            return;
        }

        foreach (var abstractNum in numbering.Elements<AbstractNum>())
        {
            if (abstractNum.AbstractNumberId?.Value is not int abstractId)
            {
                continue;
            }

            var levelMap = new Dictionary<int, NumberFormatValues>();
            foreach (var level in abstractNum.Elements<Level>())
            {
                var ilvl = level.LevelIndex?.Value ?? 0;
                var format = level.NumberingFormat?.Val?.Value ?? NumberFormatValues.Bullet;
                levelMap[ilvl] = format;
            }

            _abstractLevelFormats[abstractId] = levelMap;
        }

        foreach (var instance in numbering.Elements<NumberingInstance>())
        {
            if (instance.NumberID?.Value is not int numId)
            {
                continue;
            }

            var abstractId = instance.AbstractNumId?.Val?.Value;
            if (abstractId.HasValue)
            {
                _numIdToAbstractId[numId] = abstractId.Value;
            }
        }
    }

    /// <summary>
    /// True if <paramref name="numId"/> resolves to a known list definition;
    /// <paramref name="isOrdered"/> is then set from that <paramref name="ilvl"/>'s
    /// <see cref="NumberingFormat"/> (<c>Bullet</c>/<c>None</c> → false,
    /// anything else — Decimal, LowerLetter, UpperRoman, ... — → true, an
    /// ordered marker). Falls back to the nearest defined level if the exact
    /// <paramref name="ilvl"/> isn't defined (a foreign document's levels
    /// might not go as deep as the paragraph's own <c>ilvl</c>).
    /// </summary>
    public bool TryResolve(int numId, int ilvl, out bool isOrdered)
    {
        isOrdered = false;

        if (!_numIdToAbstractId.TryGetValue(numId, out var abstractId))
        {
            return false;
        }

        if (!_abstractLevelFormats.TryGetValue(abstractId, out var levelMap) || levelMap.Count == 0)
        {
            return false;
        }

        if (!levelMap.TryGetValue(ilvl, out var format))
        {
            var nearestLevel = levelMap.Keys.OrderBy(level => System.Math.Abs(level - ilvl)).First();
            format = levelMap[nearestLevel];
        }

        isOrdered = format != NumberFormatValues.Bullet && format != NumberFormatValues.None;
        return true;
    }
}
