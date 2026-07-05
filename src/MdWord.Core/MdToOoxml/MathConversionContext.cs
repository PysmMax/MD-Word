using System.Collections.Generic;

namespace MdWord.Core.MdToOoxml;

/// <summary>
/// Threads the per-conversion math state (Office XSL paths + the warnings
/// sink) through <see cref="BlockWalker"/>/<see cref="InlineRunBuilder"/>,
/// the same way <c>MainDocumentPart</c> is already threaded for hyperlink
/// relationships. Always non-null once it leaves <see cref="WordPackageBuilder"/>
/// — <see cref="IsActive"/> is what callers check, not nullity of this object.
/// </summary>
internal sealed class MathConversionContext
{
    public MathXslPaths XslPaths { get; }

    public List<string> Warnings { get; }

    public MathConversionContext(MathXslPaths xslPaths, List<string> warnings)
    {
        XslPaths = xslPaths;
        Warnings = warnings;
    }

    /// <summary>
    /// True when both Office XSL paths are supplied — i.e. OMML conversion
    /// should be attempted. False (xslPaths null, or either field null) means
    /// the documented "no Office XSLT available" mode: skip conversion
    /// entirely and quietly keep literal-text degradation, no warning.
    /// </summary>
    public bool IsActive => XslPaths?.Mml2OmmlXsl != null && XslPaths?.Omml2MmlXsl != null;
}
