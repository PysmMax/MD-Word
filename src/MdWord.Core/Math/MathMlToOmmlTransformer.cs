using System;
using System.Collections.Concurrent;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;

namespace MdWord.Core.Math;

/// <summary>
/// MathML → OMML via Microsoft's own <c>MML2OMML.XSL</c> (shipped with
/// Office; path supplied by the add-in as <see cref="MathXslPaths.Mml2OmmlXsl"/>).
/// <see cref="XslCompiledTransform"/> is expensive to build (compiles the
/// stylesheet), so compiled transforms are cached per XSL path — safe to
/// share across threads once built (only <see cref="XslCompiledTransform.Load(string)"/>
/// itself is not thread-safe, and that only ever runs once per path here).
/// </summary>
internal static class MathMlToOmmlTransformer
{
    private static readonly ConcurrentDictionary<string, Lazy<XslCompiledTransform>> Cache = new();

    /// <summary>
    /// Transforms a single MathML <c>&lt;math&gt;</c> element into the outer
    /// XML of an <c>m:oMath</c> fragment (no XML declaration — ready to feed
    /// straight into <c>new OfficeMath(...)</c>).
    /// </summary>
    public static string Transform(XElement mathElement, string mml2OmmlXslPath)
    {
        var transform = Cache.GetOrAdd(
            mml2OmmlXslPath,
            path => new Lazy<XslCompiledTransform>(() => LoadTransform(path))).Value;

        using var reader = mathElement.CreateReader();
        var output = new StringBuilder();
        var writerSettings = new XmlWriterSettings
        {
            OmitXmlDeclaration = true,
            ConformanceLevel = ConformanceLevel.Fragment,
        };

        using (var writer = XmlWriter.Create(output, writerSettings))
        {
            transform.Transform(reader, writer);
        }

        return output.ToString();
    }

    private static XslCompiledTransform LoadTransform(string mml2OmmlXslPath)
    {
        var xslt = new XslCompiledTransform();
        xslt.Load(mml2OmmlXslPath);
        return xslt;
    }
}
