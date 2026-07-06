using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Jint;

namespace MdWord.Core.Math;

/// <summary>
/// Lazy-singleton Jint engine that loads the vendored KaTeX bundle
/// (<c>Resources/katex.min.js</c>, embedded resource) once and exposes
/// LaTeX→MathML rendering. The first call is slow (KaTeX is a ~270KB UMD
/// bundle to parse/execute) — that is expected per the initial plan, Phase 2; a
/// background warm-up on add-in startup is Phase 6 scope, not this one.
///
/// Jint's <see cref="Engine"/> is not safe for concurrent use from multiple
/// threads, so every invocation is serialized through <see cref="EngineLock"/>.
/// </summary>
internal static class KaTeXRenderer
{
    private const string Tex2MmlDeclaration = @"
function tex2mml(tex, display) {
  return katex.renderToString(tex, { output: 'mathml', displayMode: display, throwOnError: true });
}";

    private static readonly object EngineLock = new();
    private static readonly Lazy<Engine> LazyEngine = new(CreateEngine, LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>
    /// Renders <paramref name="tex"/> to a MathML string via KaTeX. Throws
    /// (KaTeX's own <c>Jint.Runtime.JavaScriptException</c>, since
    /// <c>throwOnError: true</c>) when the LaTeX is invalid — callers are
    /// expected to catch and degrade per-formula, not let it propagate.
    /// </summary>
    public static string RenderToMathMl(string tex, bool displayMode)
    {
        var engine = LazyEngine.Value;
        lock (EngineLock)
        {
            return engine.Invoke("tex2mml", tex, displayMode).AsString();
        }
    }

    private static Engine CreateEngine()
    {
        var katexSource = LoadKatexSource();
        var engine = new Engine(options => options
            .LimitRecursion(2000)
            .TimeoutInterval(TimeSpan.FromSeconds(10))
            .LimitMemory(256 * 1024 * 1024));
        engine.Execute(katexSource);

        // Heads-up from the initial plan's Phase 2 risk list: katex.min.js is a UMD
        // bundle whose top-level `this` needs to resolve to a global object
        // for `global.katex = factory()` to stick. Verified empirically that
        // Jint 4.10.1 resolves top-level `this` to its global object, so no
        // shim is needed in practice — but check defensively and shim as a
        // fallback rather than silently building on an undefined `katex`.
        if (engine.GetValue("katex").IsUndefined())
        {
            engine.Execute("var window = this; var self = this;");
            engine.Execute(katexSource);

            if (engine.GetValue("katex").IsUndefined())
            {
                throw new ConvertException("KaTeX did not load in the JS engine (Jint): 'katex' is still undefined even after the window/self shim.");
            }
        }

        engine.Execute(Tex2MmlDeclaration);
        return engine;
    }

    private static string LoadKatexSource()
    {
        var assembly = typeof(KaTeXRenderer).Assembly;
        var resourceName = assembly.GetManifestResourceNames()
            .SingleOrDefault(name => name.EndsWith("katex.min.js", StringComparison.Ordinal));

        if (resourceName == null)
        {
            throw new ConvertException("Embedded resource 'katex.min.js' was not found in the MdWord.Core assembly.");
        }

        using var stream = assembly.GetManifestResourceStream(resourceName);
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
