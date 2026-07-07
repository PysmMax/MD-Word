using Markdig;

namespace MdWord.Core.MdToOoxml;

/// <summary>
/// Builds the single, shared Markdig pipeline configuration used across the
/// converter. Pipe tables and emphasis extras are needed by the block/inline
/// mappers landing in 1b/1c; mathematics is parsed starting now (per the
/// initial plan §4) but not converted to OMML until Phase 2 — until then math nodes are
/// simply not mapped by the (currently paragraph-only) block walker.
/// </summary>
internal static class MarkdigPipelineFactory
{
    public static MarkdownPipeline Create()
    {
        return new MarkdownPipelineBuilder()
            .UsePipeTables()
            // All extras: since v1.0.2 the inline mapper has real OOXML
            // mappings for subscript (~x~ -> w:vertAlign), superscript
            // (^x^ -> w:vertAlign), inserted (++x++ -> w:u) and marked
            // (==x== -> w:highlight), alongside strikethrough (~~x~~).
            .UseEmphasisExtras(Markdig.Extensions.EmphasisExtras.EmphasisExtraOptions.Default)
            .UseMathematics()
            .Build();
    }
}
