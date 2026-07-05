using System;
using System.IO;

namespace MdWord.Core.Tests;

/// <summary>
/// Single place the math/round-trip tests resolve Office's
/// MML2OMML.XSL / OMML2MML.XSL from, instead of three hardcoded copies of
/// one machine's path. Order: MDWORD_OFFICE_XSL_DIR env var (explicit
/// override, e.g. CI), then the standard Click-to-Run locations. When
/// nothing matches, falls back to the historical hardcoded path so the
/// failure mode stays the same loud one (file not found -> degrade ->
/// assertion fails), never a silent skip.
/// </summary>
internal static class OfficeXslLocator
{
    public static MathXslPaths Resolve()
    {
        foreach (var dir in CandidateDirectories())
        {
            if (string.IsNullOrEmpty(dir))
            {
                continue;
            }

            var mml2Omml = Path.Combine(dir, "MML2OMML.XSL");
            var omml2Mml = Path.Combine(dir, "OMML2MML.XSL");
            if (File.Exists(mml2Omml) && File.Exists(omml2Mml))
            {
                return new MathXslPaths { Mml2OmmlXsl = mml2Omml, Omml2MmlXsl = omml2Mml };
            }
        }

        return new MathXslPaths
        {
            Mml2OmmlXsl = @"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
            Omml2MmlXsl = @"C:\Program Files\Microsoft Office\root\Office16\OMML2MML.XSL",
        };
    }

    private static string[] CandidateDirectories() => new[]
    {
        Environment.GetEnvironmentVariable("MDWORD_OFFICE_XSL_DIR"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Microsoft Office\root\Office16"),
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft Office\root\Office16"),
    };
}
