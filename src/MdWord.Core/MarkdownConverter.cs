using System;
using DocumentFormat.OpenXml.Packaging;
using Markdig;
using MdWord.Core.MdToOoxml;
using MdWord.Core.OoxmlToMd;

namespace MdWord.Core;

/// <summary>
/// Paths to Office's own MML2OMML.XSL / OMML2MML.XSL, supplied by the add-in
/// (resolved from the Word installation directory). When this is null, or
/// either field is null, formula conversion degrades to literal text —
/// documented, intentional "no Office XSLT available" behavior (Phase 2),
/// not a bug. Both fields must be non-null for OMML conversion to activate,
/// even though only <see cref="Mml2OmmlXsl"/> is used until Phase 4 wires up
/// the OMML→MathML direction — an incomplete pair is treated the same as
/// "not supplied".
/// </summary>
public sealed class MathXslPaths
{
    public string Mml2OmmlXsl;
    public string Omml2MmlXsl;
}

/// <summary>Result of <see cref="MarkdownConverter.ToOoxml"/>.</summary>
public sealed class OoxmlResult
{
    /// <summary>Flat OPC XML string — for Range.InsertXML.</summary>
    public string FlatOpc;

    /// <summary>Raw docx bytes — for the Range.InsertFile fallback.</summary>
    public byte[] DocxBytes;

    /// <summary>Human-readable notes about degraded elements (e.g. "formula X inserted as text").</summary>
    public string[] Warnings;
}

/// <summary>Result of <see cref="MarkdownConverter.ToMarkdown"/>.</summary>
public sealed class MarkdownResult
{
    public string Markdown;
    public string[] Warnings;
}

/// <summary>
/// Public facade of MdWord.Core: Markdig ⇄ WordprocessingML conversion,
/// independent of Word. See the initial plan §6 for the full core contract.
/// </summary>
public sealed class MarkdownConverter
{
    private static readonly MarkdownPipeline Pipeline = MarkdigPipelineFactory.Create();

    private readonly MathXslPaths _xslPaths;

    public MarkdownConverter(MathXslPaths xslPaths)
    {
        _xslPaths = xslPaths;
    }

    /// <summary>
    /// Pre-parses and executes the KaTeX bundle inside Jint (the ~1s
    /// first-use cost) so the first real formula click doesn't pay it on
    /// Word's UI thread. Safe to call from a background thread — the engine
    /// is a lazy singleton serialized by its own lock. No Office XSL paths
    /// needed: warm-up stops at the LaTeX→MathML stage.
    /// </summary>
    public static void WarmUpMathEngine()
    {
        Math.KaTeXRenderer.RenderToMathMl("x", displayMode: false);
    }

    /// <summary>Markdown → WordprocessingML (Phases 1–2). Never throws on a single bad element.</summary>
    public OoxmlResult ToOoxml(string markdown)
    {
        if (markdown == null)
        {
            throw new ConvertException("Input Markdown text cannot be null.");
        }

        // AI tools often emit \(...\)/\[...\] instead of $...$/$$...$$; rewrite
        // before parsing so Markdig's mathematics extension actually sees math
        // nodes (the initial plan, Phase 2). Line-oriented, skips fenced code blocks.
        var preprocessed = Math.LatexDelimiterPreprocessor.Rewrite(markdown);

        Markdig.Syntax.MarkdownDocument document;
        try
        {
            document = Markdown.Parse(preprocessed, Pipeline);
        }
        catch (Exception ex)
        {
            throw new ConvertException("Could not parse the Markdown.", ex);
        }

        byte[] docxBytes;
        string flatOpc;
        string[] warnings;
        try
        {
            (docxBytes, flatOpc, warnings) = WordPackageBuilder.Build(document, _xslPaths);
        }
        catch (Exception ex)
        {
            throw new ConvertException("Could not generate the Word document.", ex);
        }

        return new OoxmlResult
        {
            DocxBytes = docxBytes,
            FlatOpc = flatOpc,
            Warnings = warnings,
        };
    }

    /// <summary>
    /// Flat OPC WordprocessingML → Markdown (Phase 4). Never throws on a
    /// single degraded element — those collect into <see cref="MarkdownResult.Warnings"/>
    /// instead (mirrors <see cref="ToOoxml"/>'s per-element-failure contract).
    /// </summary>
    public MarkdownResult ToMarkdown(string flatOpcXml)
    {
        if (flatOpcXml == null)
        {
            throw new ConvertException("Input Flat OPC XML cannot be null.");
        }

        WordprocessingDocument document;
        try
        {
            document = WordprocessingDocument.FromFlatOpcString(flatOpcXml);
        }
        catch (Exception ex)
        {
            throw new ConvertException("Could not parse the Flat OPC XML.", ex);
        }

        using (document)
        {
            string markdown;
            string[] warnings;
            try
            {
                (markdown, warnings) = DocumentMarkdownBuilder.Build(document, _xslPaths);
            }
            catch (Exception ex)
            {
                throw new ConvertException("Could not convert the Word document to Markdown.", ex);
            }

            return new MarkdownResult { Markdown = markdown, Warnings = warnings };
        }
    }
}
