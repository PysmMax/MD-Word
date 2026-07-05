using System;

namespace MdWord.Core;

/// <summary>
/// Thrown when a document-level conversion failure occurs (e.g. the input
/// cannot be parsed or the resulting document cannot be generated at all).
/// Failures scoped to a single element must NOT throw this — they degrade
/// to literal text plus a warning instead (see <see cref="OoxmlResult.Warnings"/>).
/// </summary>
public sealed class ConvertException : Exception
{
    public ConvertException(string message)
        : base(message)
    {
    }

    public ConvertException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
