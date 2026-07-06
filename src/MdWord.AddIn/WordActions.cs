using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MdWord.Core;

namespace MdWord.AddIn;

/// <summary>
/// The actual Word-object-model orchestration for "Insert Markdown"
/// (initial plan, Phase 3). Deliberately kept out of <see cref="Connect"/> so that
/// every reference to <c>MdWord.Core</c> types lives in one place, JIT'd only
/// when this class's methods are first entered — by then
/// <see cref="Connect"/>'s static constructor has already installed the
/// <see cref="AppDomain.AssemblyResolve"/> handler that resolves
/// <c>MdWord.Core.dll</c> and its own dependencies (Markdig,
/// DocumentFormat.OpenXml.*, Jint, System.IO.Packaging) from the add-in's own
/// bin folder — Word's <c>mscoree.dll</c>/<c>CodeBase</c> activation only
/// probes for <c>MdWord.AddIn.dll</c> itself, not its dependency graph.
/// </summary>
internal static class WordActions
{
    private const string UndoRecordName = "Insert Markdown";

    /// <summary>Background warm-up passthrough (see Connect.OnStartupComplete).</summary>
    public static void WarmUpMathEngine()
    {
        try
        {
            MarkdownConverter.WarmUpMathEngine();
        }
        catch (Exception ex)
        {
            // Warm-up is best-effort: a failure here must never surface UI —
            // the real conversion path has its own degrade-and-warn handling.
            Logger.LogWarning($"Background KaTeX warm-up failed ({ex.Message}).");
        }
    }

    /// <summary>
    /// Clipboard → Markdown → WordprocessingML → inserted at the current
    /// selection, as a single undo step. <paramref name="wordApp"/> is the
    /// live <c>Word.Application</c> COM object, passed as <c>dynamic</c> so
    /// this project needs no compile-time dependency beyond what's already
    /// referenced (Interop types are still embedded via the csproj's
    /// <c>EmbedInteropTypes</c>, but calls here are late-bound to avoid
    /// pulling extra typed surface into this file).
    /// </summary>
    public static void PasteMarkdown(dynamic wordApp)
    {
        if (IsPasteBlockedByProtectedViewOrReadOnly(wordApp))
        {
            MessageBox.Show(
                "The document is open in Protected View or read-only, so " +
                "Markdown cannot be inserted. Enable editing (or close " +
                "Protected View) and try again.",
                "MD-Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        string markdown = Clipboard.GetText();

        if (string.IsNullOrWhiteSpace(markdown))
        {
            MessageBox.Show(
                "The clipboard is empty or contains no text.",
                "MD-Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        const int MaxInputChars = 1_000_000; // Per the initial plan, Phase 6: warn on >1 MB of text
        if (markdown.Length > MaxInputChars)
        {
            MessageBox.Show(
                $"The clipboard has too much text ({markdown.Length:N0} characters; maximum {MaxInputChars:N0}). " +
                "Paste a smaller fragment.",
                "MD-Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        var xslPaths = ResolveXslPaths(wordApp);
        var result = new MarkdownConverter(xslPaths).ToOoxml(markdown);

        wordApp.UndoRecord.StartCustomRecord(UndoRecordName);
        try
        {
            InsertResult(wordApp, result);

            if (result.Warnings != null && result.Warnings.Length > 0)
            {
                wordApp.StatusBar = string.Join("; ", result.Warnings);
            }
        }
        finally
        {
            wordApp.UndoRecord.EndCustomRecord();
        }
    }

    /// <summary>
    /// Primary path: <c>Range.InsertXML</c> on the Flat OPC string. Falls
    /// back to writing <see cref="OoxmlResult.DocxBytes"/> to a temp .docx
    /// and <c>Range.InsertFile</c> on <see cref="COMException"/> — per
    /// the initial plan, Phase 3's documented fallback. Step 0's WebFetch of the
    /// official <c>Range.InsertXML</c>/<c>Range.WordOpenXML</c> VBA reference
    /// pages did not confirm or deny Flat-OPC-specific behavior (the
    /// InsertXML page's own example is generic custom XML, not
    /// WordprocessingML), so this fallback is kept exactly as designed rather
    /// than removed or made primary — only the live checkpoint can confirm
    /// which path Word actually needs.
    /// </summary>
    private static void InsertResult(dynamic wordApp, OoxmlResult result)
    {
        try
        {
            wordApp.Selection.Range.InsertXML(result.FlatOpc);
        }
        catch (COMException ex)
        {
            Logger.LogWarning($"Range.InsertXML failed ({ex.Message}), trying InsertFile.");
            InsertViaTempFile(wordApp, result);
        }
    }

    private static void InsertViaTempFile(dynamic wordApp, OoxmlResult result)
    {
        // Random name: two concurrent Word processes must not fight over one
        // fixed temp file in the InsertFile fallback path.
        string tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        File.WriteAllBytes(tempPath, result.DocxBytes);
        try
        {
            wordApp.Selection.Range.InsertFile(tempPath);
        }
        finally
        {
            try
            {
                File.Delete(tempPath);
            }
            catch
            {
                // Best-effort cleanup; a leftover temp file is not worth failing the paste over.
            }
        }
    }

    /// <summary>
    /// Word selection → Markdown → clipboard (the initial plan, Phase 4). Mirror of
    /// <see cref="PasteMarkdown"/>'s structure: empty-selection no-op,
    /// resolve XSL paths, delegate to <c>MdWord.Core</c>, surface warnings
    /// via StatusBar. Unlike the paste path there is no Word-document
    /// mutation here (no undo record needed) — copying to the clipboard is
    /// not an undoable document edit.
    /// </summary>
    public static void CopyMarkdown(dynamic wordApp)
    {
        bool isEmpty;
        try
        {
            isEmpty = (int)wordApp.Selection.Type == WdSelectionTypeNoSelection
                || string.IsNullOrWhiteSpace((string)wordApp.Selection.Range.Text);
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Could not read the selection ({ex.Message}) — treating it as empty.");
            isEmpty = true;
        }

        if (isEmpty)
        {
            MessageBox.Show(
                "Select text in the document before copying it as Markdown.",
                "MD-Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        const int MaxSelectionChars = 1_000_000; // Per the initial plan, Phase 6: warn on >1 MB of text
        int selectionLength = 0;
        try
        {
            selectionLength = (int)wordApp.Selection.Range.End - (int)wordApp.Selection.Range.Start;
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Could not measure the selection size ({ex.Message}) — continuing without a limit.");
        }

        if (selectionLength > MaxSelectionChars)
        {
            MessageBox.Show(
                $"The selection is too large ({selectionLength:N0} characters; maximum {MaxSelectionChars:N0}). " +
                "Select and copy a smaller fragment.",
                "MD-Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        var xslPaths = ResolveXslPaths(wordApp);
        string flatOpc = wordApp.Selection.Range.WordOpenXML;
        var result = new MarkdownConverter(xslPaths).ToMarkdown(flatOpc);

        if (string.IsNullOrEmpty(result.Markdown))
        {
            // A non-empty selection can still yield empty Markdown (e.g. it
            // contained only unsupported content). Clipboard.SetText throws
            // ArgumentException on an empty string, so guard explicitly
            // rather than let that escape as a confusing error dialog.
            MessageBox.Show(
                "Could not produce Markdown from the selection (it may contain only unsupported content).",
                "MD-Word",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        // SetText throws ExternalException when another process momentarily
        // holds the clipboard (clipboard managers, RDP) -- retry briefly.
        Clipboard.SetDataObject(result.Markdown, true, 5, 100);

        if (result.Warnings != null && result.Warnings.Length > 0)
        {
            wordApp.StatusBar = string.Join("; ", result.Warnings);
        }
    }

    /// <summary>
    /// <c>WdSelectionType.wdNoSelection</c>'s integer value (0) — used here
    /// instead of the typed Interop enum so this file's late-bound
    /// <c>dynamic</c> style (see the class-level remarks) stays consistent;
    /// avoids adding an early-bound Interop type reference just for one
    /// constant.
    /// </summary>
    private const int WdSelectionTypeNoSelection = 0;

    /// <summary>
    /// Guards <see cref="PasteMarkdown"/> against two states where
    /// <c>Range.InsertXML</c>/<c>InsertFile</c> would otherwise fail with a
    /// confusing COM error: a document open in Protected View (email
    /// attachments, downloaded files), or a plain read-only document. Fails
    /// open (returns false, letting the paste attempt proceed) if the check
    /// itself can't be made -- a missed guard just falls through to the
    /// normal COM-failure handling in <see cref="Connect.HandleFailure"/>.
    /// </summary>
    private static bool IsPasteBlockedByProtectedViewOrReadOnly(dynamic wordApp)
    {
        try
        {
            if (wordApp.ActiveProtectedViewWindow != null)
            {
                return true;
            }
        }
        catch
        {
            // ActiveProtectedViewWindow unavailable on this Word version -- fall through.
        }

        try
        {
            return (bool)wordApp.ActiveDocument.ReadOnly;
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Could not check Protected View/read-only state ({ex.Message}) — continuing.");
            return false;
        }
    }

    /// <summary>
    /// Resolves Office's own MML2OMML.XSL/OMML2MML.XSL from the running
    /// Word installation directory (<c>wordApp.Path</c>). Degrades to
    /// <c>null</c> (formulas become literal text, per
    /// <see cref="MathXslPaths"/>'s documented contract) rather than
    /// throwing when the files aren't found at an unusual install layout —
    /// logs a warning instead.
    /// </summary>
    private static MathXslPaths ResolveXslPaths(dynamic wordApp)
    {
        string wordPath;
        try
        {
            wordPath = (string)wordApp.Path;
        }
        catch (Exception ex)
        {
            Logger.LogWarning($"Could not read wordApp.Path ({ex.Message}) — formulas will be plain text.");
            return null;
        }

        if (string.IsNullOrEmpty(wordPath))
        {
            return null;
        }

        string mml2Omml = Path.Combine(wordPath, "MML2OMML.XSL");
        string omml2Mml = Path.Combine(wordPath, "OMML2MML.XSL");

        if (!File.Exists(mml2Omml) || !File.Exists(omml2Mml))
        {
            Logger.LogWarning(
                $"MML2OMML.XSL/OMML2MML.XSL not found in '{wordPath}' — formulas will be inserted as plain text.");
            return null;
        }

        return new MathXslPaths { Mml2OmmlXsl = mml2Omml, Omml2MmlXsl = omml2Mml };
    }
}
