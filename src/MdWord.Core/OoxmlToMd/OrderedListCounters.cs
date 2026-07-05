using System.Collections.Generic;
using System.Linq;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Tracks per-<c>ilvl</c> ordered-list counters while walking a contiguous
/// run of list paragraphs. Word stores lists as a flat sequence of
/// paragraphs (each carrying its own <c>numId</c>/<c>ilvl</c>), not a nested
/// block structure, so nested-list numbering has to be reconstructed here
/// rather than inherited from a parent block. Resets deeper levels whenever
/// the walker returns to a shallower level (so a later re-descent restarts
/// at 1) and resets everything whenever the containing <c>numId</c> changes
/// (a different list instance) or the list run is interrupted entirely.
/// </summary>
internal sealed class OrderedListCounters
{
    private readonly Dictionary<int, int> _counts = new();
    private int? _lastNumId;

    public int Next(int numId, int ilvl)
    {
        if (_lastNumId != numId)
        {
            _counts.Clear();
            _lastNumId = numId;
        }

        foreach (var deeperLevel in _counts.Keys.Where(level => level > ilvl).ToList())
        {
            _counts.Remove(deeperLevel);
        }

        var next = _counts.TryGetValue(ilvl, out var current) ? current + 1 : 1;
        _counts[ilvl] = next;
        return next;
    }

    public void Reset()
    {
        _counts.Clear();
        _lastNumId = null;
    }
}
