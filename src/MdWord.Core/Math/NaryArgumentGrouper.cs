using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace MdWord.Core.Math;

/// <summary>
/// MML2OMML.XSL's msub/msup/msubsup/munder/mover/munderover templates build
/// an n-ary operator's argument (<c>&lt;m:e&gt;</c>) by applying
/// <c>NaryHandleMrowMstyle</c> to <c>following-sibling::*[1]</c> — and that
/// template only pulls content in when that single following sibling is
/// itself an <c>mrow</c> or <c>mstyle</c> (confirmed by reading MML2OMML.XSL
/// directly). KaTeX's MathML output never wraps an n-ary's argument that
/// way — it emits flat siblings (a bare <c>msup</c>, <c>mi</c>,
/// <c>mtext</c>, ...) — so the argument is silently dropped: Word renders an
/// empty dashed placeholder box where the argument should be, with the real
/// argument left as disconnected text after it. Converting the pasted
/// equation to Word's linear format and back "fixes" it only because that
/// round-trip goes through Word's own (different) equation parser, not this
/// XSLT.
///
/// This groups the following siblings that belong to an n-ary's argument
/// into an explicit <c>mrow</c> before the MathML reaches
/// <see cref="MathMlToOmmlTransformer"/>, stopping before a low-precedence
/// relational operator so <c>∫ f(x) dx = g</c> still separates "= g" from
/// the integral (matching what MML2OMML.XSL does correctly for an n-ary
/// argument that already arrives pre-grouped as an mrow).
/// </summary>
internal static class NaryArgumentGrouper
{
    private static readonly XNamespace Mml = "http://www.w3.org/1998/Math/MathML";

    // Exactly MML2OMML.XSL's own isNaryOper character set (line ~3639) --
    // wrapping only kicks in for the same operators the XSLT itself treats
    // as n-ary, so this stays a no-op everywhere else.
    private static readonly HashSet<string> NaryChars = new()
    {
        "∫", "∬", "∭", "∮", "∯", "∰", "∲", "∳", "∱",
        "∩", "∪", "∏", "∐", "∑", "⋀", "⋁", "⋂", "⋃",
    };

    // Narrower than NaryChars: ∩/∪ (U+2229/U+222A, \cap/\cup) are *binary*
    // infix operators when bare -- unlike \int/\sum/\bigcap, synthesizing a
    // prefix n-ary wrapper for a bare one would orphan its left operand and
    // wrongly swallow its right operand as the sole argument. Only
    // genuinely prefix operators are safe to bare-wrap; the "big" set
    // variants (⋂/⋃, \bigcap/\bigcup) are separate codepoints and stay in.
    // This set only gates WrapBareNaryOperators -- an already-limited
    // msub/msup/... around ∩/∪ (e.g. "\cap_a^b", however unusual) still
    // matches MML2OMML.XSL's own isNaryOper via IsNaryWrapper, unchanged.
    private static readonly HashSet<string> BareWrappableNaryChars = new(NaryChars.Except(new[] { "∩", "∪" }));

    private static readonly HashSet<string> NaryWrapperElements = new()
    {
        "msub", "msup", "msubsup", "munder", "mover", "munderover",
    };

    private static readonly HashSet<string> RelationalOperators = new()
    {
        "=", "<", ">", "≤", "≥", "≠", "≈", "→",
    };

    /// <summary>Mutates <paramref name="mathElement"/> in place.</summary>
    public static void GroupArguments(XElement mathElement)
    {
        // A bare n-ary operator with no {}^{} limits (e.g. plain "\int f dx")
        // has no msub/msup/... wrapper at all -- MML2OMML.XSL has no
        // template for a standalone mml:mo, so it would otherwise fall back
        // to flat literal text with no <m:nary> whatsoever. Synthesize an
        // msup wrapper (with an empty, placeholder superscript) so the
        // normal msup template fires; MathConverter hides the resulting
        // empty limit slot after the transform (see
        // NaryHiddenLimitFixup) since MML2OMML.XSL's own subHide/supHide
        // table has no combination reachable from input shape alone that
        // hides both limits together.
        WrapBareNaryOperators(mathElement);

        // Snapshot first: grouping mutates sibling structure while walking.
        var naryWrappers = mathElement.Descendants()
            .Where(e => NaryWrapperElements.Contains(e.Name.LocalName) && IsNaryWrapper(e))
            .ToList();

        foreach (var wrapper in naryWrappers)
        {
            GroupFollowingSiblingsIfNeeded(wrapper);
        }
    }

    private static void WrapBareNaryOperators(XElement mathElement)
    {
        var bareOperators = mathElement.Descendants(Mml + "mo")
            .Where(mo => BareWrappableNaryChars.Contains(mo.Value.Trim())
                         && !NaryWrapperElements.Contains(mo.Parent?.Name.LocalName))
            .ToList();

        foreach (var mo in bareOperators)
        {
            var wrapper = new XElement(Mml + "msup", new XElement(Mml + "mrow"));
            mo.ReplaceWith(wrapper);
            wrapper.AddFirst(mo);
        }
    }

    private static bool IsNaryWrapper(XElement element)
    {
        var firstChild = element.Elements().FirstOrDefault();
        return firstChild != null && NaryChars.Contains(firstChild.Value.Trim());
    }

    private static void GroupFollowingSiblingsIfNeeded(XElement wrapper)
    {
        var next = wrapper.ElementsAfterSelf().FirstOrDefault();
        if (next == null)
        {
            return; // No argument at all -- leave <m:e/> empty, as before.
        }

        if (next.Name.LocalName is "mrow" or "mstyle")
        {
            return; // Already the shape MML2OMML.XSL expects.
        }

        var toAbsorb = new List<XElement>();
        foreach (var sibling in wrapper.ElementsAfterSelf())
        {
            if (sibling.Name.LocalName == "mo" && RelationalOperators.Contains(sibling.Value.Trim()))
            {
                break;
            }

            toAbsorb.Add(sibling);
        }

        if (toAbsorb.Count == 0)
        {
            return;
        }

        var group = new XElement(Mml + "mrow");
        wrapper.AddAfterSelf(group);
        foreach (var element in toAbsorb)
        {
            element.Remove();
            group.Add(element);
        }
    }
}
