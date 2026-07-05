using Markdig;

namespace MdWord.Core.MdToOoxml;

/// <summary>
/// Builds the single, shared Markdig pipeline configuration used across the
/// converter. Pipe tables and emphasis extras are needed by the block/inline
/// mappers landing in 1b/1c; mathematics is parsed starting now (per PLAN.md
/// §4) but not converted to OMML until Phase 2 — until then math nodes are
/// simply not mapped by the (currently paragraph-only) block walker.
/// </summary>
internal static class MarkdigPipelineFactory
{
    public static MarkdownPipeline Create()
    {
        return new MarkdownPipelineBuilder()
            .UsePipeTables()
            // Only strikethrough: the inline mapper has no OOXML mapping for
            // subscript (~x~), superscript (^x^), inserted (++x++) or marked
            // (==x==) -- with the default (all-on) options those silently came
            // out as strike/italic/bold. Literal text is the honest degradation.
            .UseEmphasisExtras(Markdig.Extensions.EmphasisExtras.EmphasisExtraOptions.Strikethrough)
            .UseMathematics()
            .Build();
    }
}
