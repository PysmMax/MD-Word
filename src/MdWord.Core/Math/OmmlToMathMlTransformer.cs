using System;
using System.Collections.Concurrent;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;

namespace MdWord.Core.Math;

/// <summary>
/// OMML → MathML via Microsoft's own <c>OMML2MML.XSL</c> (shipped with
/// Office; path supplied by the add-in as <see cref="MathXslPaths.Omml2MmlXsl"/>).
/// Sibling of <see cref="MathMlToOmmlTransformer"/> (forward direction) —
/// same cached-<see cref="XslCompiledTransform"/>-per-path pattern, kept as
/// a separate class rather than bolted onto the forward one (Phase 4 brief).
/// </summary>
internal static class OmmlToMathMlTransformer
{
    private static readonly ConcurrentDictionary<string, Lazy<XslCompiledTransform>> Cache = new();

    /// <summary>
    /// Transforms a single OMML <c>m:oMath</c> element into the MathML
    /// <c>&lt;math&gt;</c> element produced by Office's own OMML2MML.XSL.
    /// </summary>
    public static XElement Transform(XElement ommlElement, string omml2MmlXslPath)
    {
        var transform = Cache.GetOrAdd(
            omml2MmlXslPath,
            path => new Lazy<XslCompiledTransform>(() => LoadTransform(path))).Value;

        using var reader = ommlElement.CreateReader();
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

        return XElement.Parse(output.ToString());
    }

    private static XslCompiledTransform LoadTransform(string omml2MmlXslPath)
    {
        var xslt = new XslCompiledTransform();
        xslt.Load(omml2MmlXslPath);
        return xslt;
    }
}
