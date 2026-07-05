using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MdWord.Core.OoxmlToMd;

/// <summary>
/// Maps a <c>w:sym</c> (SymbolChar) that uses the legacy Adobe "Symbol" font
/// to its Unicode equivalent. Word stores e.g. a delta as
/// <c>&lt;w:sym w:font="Symbol" w:char="F064"/&gt;</c>; the low byte (0x64)
/// is the Symbol-font code point. Without this, such characters were dropped
/// silently on the copy path (LIVE-5). Covers the Greek alphabet (the common
/// real-world case — physics/maths variables); anything unmapped degrades to
/// a Warning so nothing is lost silently.
/// </summary>
internal static class SymbolFontMap
{
    private static readonly Dictionary<int, char> SymbolToUnicode = new()
    {
        // lowercase Greek
        [0x61] = 'α', [0x62] = 'β', [0x67] = 'γ', [0x64] = 'δ', [0x65] = 'ε',
        [0x7A] = 'ζ', [0x68] = 'η', [0x71] = 'θ', [0x69] = 'ι', [0x6B] = 'κ',
        [0x6C] = 'λ', [0x6D] = 'μ', [0x6E] = 'ν', [0x78] = 'ξ', [0x6F] = 'ο',
        [0x70] = 'π', [0x72] = 'ρ', [0x73] = 'σ', [0x74] = 'τ', [0x75] = 'υ',
        [0x66] = 'φ', [0x63] = 'χ', [0x79] = 'ψ', [0x77] = 'ω',
        // uppercase Greek
        [0x41] = 'Α', [0x42] = 'Β', [0x47] = 'Γ', [0x44] = 'Δ', [0x45] = 'Ε',
        [0x5A] = 'Ζ', [0x48] = 'Η', [0x51] = 'Θ', [0x49] = 'Ι', [0x4B] = 'Κ',
        [0x4C] = 'Λ', [0x4D] = 'Μ', [0x4E] = 'Ν', [0x58] = 'Ξ', [0x4F] = 'Ο',
        [0x50] = 'Π', [0x52] = 'Ρ', [0x53] = 'Σ', [0x54] = 'Τ', [0x55] = 'Υ',
        [0x46] = 'Φ', [0x43] = 'Χ', [0x59] = 'Ψ', [0x57] = 'Ω',
    };

    public static string Resolve(SymbolChar sym, MdConversionContext context)
    {
        var font = sym.Font?.Value;
        var raw = sym.Char?.Value;

        if (string.Equals(font, "Symbol", System.StringComparison.OrdinalIgnoreCase)
            && raw != null
            && int.TryParse(raw, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var code)
            && SymbolToUnicode.TryGetValue(code & 0xFF, out var mapped))
        {
            return mapped.ToString();
        }

        context.Warnings.Add($"Символ (шрифт {font ?? "?"}, код {raw ?? "?"}) пропущено — не в таблиці SymbolFontMap.");
        return string.Empty;
    }
}
