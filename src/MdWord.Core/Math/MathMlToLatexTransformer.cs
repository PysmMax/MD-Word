using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using System.Xml.Xsl;

namespace MdWord.Core.Math;

/// <summary>
/// MathML → LaTeX via the vendored xsltml 2.1.2 library (<c>mmltex.xsl</c> +
/// its includes: tokens/glayout/scripts/tables/entities/cmarkup — see
/// <c>Resources/mmltex/README-xsltml.txt</c> for the license, and
/// <c>THIRD-PARTY-NOTICES.md</c> for the source/version/license summary). Unlike
/// <see cref="MathMlToOmmlTransformer"/>/<see cref="OmmlToMathMlTransformer"/>,
/// this transform is entirely self-contained (no externally-supplied path):
/// the library ships as embedded resources and is extracted to a temp
/// directory once per process, then loaded from that physical path —
/// <see cref="XslCompiledTransform"/>'s <c>xsl:include</c>/<c>xsl:import</c>
/// resolution needs real files on disk, and a custom <see cref="System.Xml.XmlResolver"/>
/// mapping embedded-resource streams is more code for no benefit here (per
/// the Phase 4 brief's explicit guidance).
///
/// The entry point loaded is <c>mdword-entry.xsl</c> (our own file, not part
/// of the vendored xsltml distribution) rather than <c>mmltex.xsl</c>
/// directly — it imports mmltex.xsl unmodified but overrides the top-level
/// <c>m:math</c> template so this transform's output is the bare LaTeX body,
/// with no <c>$...$</c>/<c>\[...\]</c> wrapping. MD-Word's own Markdown
/// delimiter convention (<c>$...$</c> inline, <c>$$...$$</c> display) is
/// applied by the caller in <c>MdWord.Core.OoxmlToMd</c>, which already
/// knows whether it is converting an inline <c>m:oMath</c> or a block
/// <c>m:oMathPara</c> — confirmed empirically that mmltex.xsl's own
/// delimiters would otherwise double up with ours.
/// </summary>
internal static class MathMlToLatexTransformer
{
    private const string EntryResourceFileName = "mdword-entry.xsl";

    private static readonly string[] VendoredResourceFileNames =
    {
        "tokens.xsl",
        "glayout.xsl",
        "scripts.xsl",
        "tables.xsl",
        "entities.xsl",
        "cmarkup.xsl",
        "mmltex.xsl",
        EntryResourceFileName,
    };

    // ExecutionAndPublication (not PublicationOnly): PublicationOnly lets
    // multiple threads run ExtractEmbeddedXslSet concurrently, and that method
    // writes the SAME fixed temp files via File.Create -- racing threads hit
    // real file-sharing violations against each other and intermittently fail
    // (~35% of full test-suite runs under xUnit's parallel execution).
    // ExecutionAndPublication serializes the factory in-process, which
    // eliminates that race. It does mean one transient failure (e.g. genuine
    // cross-process contention from two separate Word.exe instances) stays
    // cached for the process lifetime -- accepted as a rare, low-severity
    // trade-off (REL-13) rather than risk the proven in-process race.
    private static readonly Lazy<XslCompiledTransform> LazyTransform =
        new(LoadTransform, LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>
    /// Transforms a single MathML <c>&lt;math&gt;</c> element into its bare
    /// LaTeX body (no surrounding Markdown/LaTeX math delimiters).
    /// </summary>
    public static string Transform(XElement mathElement)
    {
        var transform = LazyTransform.Value;

        using var reader = mathElement.CreateReader();
        var output = new StringWriter();

        // mmltex.xsl declares <xsl:output method="text"/> — using a plain
        // TextWriter (not an XmlWriter) here avoids XML-entity-escaping
        // characters like '<'/'&' that can legitimately appear in emitted
        // LaTeX (e.g. \binom, comparison operators inside \text{}).
        transform.Transform(reader, null, output);

        return output.ToString();
    }

    private static XslCompiledTransform LoadTransform()
    {
        var entryPath = ExtractEmbeddedXslSet();
        var xslt = new XslCompiledTransform();
        xslt.Load(entryPath);
        return xslt;
    }

    /// <summary>
    /// Writes the embedded xsltml files (+ our own entry stylesheet) to a
    /// per-process temp directory so <see cref="XslCompiledTransform.Load(string)"/>
    /// can resolve the <c>xsl:include</c>/<c>xsl:import</c> chain via normal
    /// filesystem-relative URI resolution. Runs once per process (guarded by
    /// <see cref="LazyTransform"/>); re-extracting on every call would be
    /// wasteful but harmless (idempotent overwrite of the same content).
    /// </summary>
    private static string ExtractEmbeddedXslSet()
    {
        var assembly = typeof(MathMlToLatexTransformer).Assembly;
        var resourceNames = assembly.GetManifestResourceNames();

        var targetDir = Path.Combine(Path.GetTempPath(), "MdWord", "mmltex");
        Directory.CreateDirectory(targetDir);

        foreach (var fileName in VendoredResourceFileNames)
        {
            var resourceName = resourceNames.SingleOrDefault(
                name => name.EndsWith("." + fileName, StringComparison.Ordinal));

            if (resourceName == null)
            {
                throw new ConvertException($"Embedded resource '{fileName}' (mmltex) was not found in the MdWord.Core assembly.");
            }

            var targetPath = Path.Combine(targetDir, fileName);
            using var resourceStream = assembly.GetManifestResourceStream(resourceName);
            using var fileStream = File.Create(targetPath);
            resourceStream!.CopyTo(fileStream);
        }

        return Path.Combine(targetDir, EntryResourceFileName);
    }
}
