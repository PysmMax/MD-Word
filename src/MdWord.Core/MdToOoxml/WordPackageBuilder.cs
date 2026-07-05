using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml; // WordprocessingDocumentType (root namespace in SDK 3.x)
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace MdWord.Core.MdToOoxml;

/// <summary>
/// Scaffolds a docx package (MainDocumentPart + minimal StyleDefinitionsPart)
/// in a <see cref="MemoryStream"/> via the OpenXML SDK, and returns both raw
/// docx bytes and the Flat OPC XML string for the same content. No
/// <see cref="SectionProperties"/> is emitted, so inserting the result into
/// an existing Word document does not drag in a section break.
/// Block/inline content mapping itself lives in <see cref="BlockWalker"/>/
/// <see cref="InlineRunBuilder"/>.
/// </summary>
internal static class WordPackageBuilder
{
    public static (byte[] DocxBytes, string FlatOpc, string[] Warnings) Build(MarkdownDocument markdownDocument, MathXslPaths xslPaths)
    {
        using var stream = new MemoryStream();
        string flatOpc;
        var warnings = new List<string>();
        var mathContext = new MathConversionContext(xslPaths, warnings);

        using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDocument.AddMainDocumentPart();

            var body = new Body();
            foreach (var element in BlockWalker.MapBlocks(markdownDocument, mainPart, mathContext))
            {
                body.Append(element);
            }

            mainPart.Document = new Document(body);

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = StylesPartBuilder.BuildMinimalStyles();
            stylesPart.Styles.Save();

            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = NumberingPartBuilder.BuildMinimalNumbering();
            numberingPart.Numbering.Save();

            mainPart.Document.Save();

            flatOpc = wordDocument.ToFlatOpcString();
        }

        return (stream.ToArray(), flatOpc, warnings.ToArray());
    }
}
