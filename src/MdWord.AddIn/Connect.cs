using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MdWord.AddIn;

/// <summary>
/// COM entry point for the add-in. Registered under the fixed
/// <c>CLSID_Connect</c> (also baked into installer/MdWord.iss's
/// [Registry] section).
/// </summary>
/// <remarks>
/// <para>
/// <b>Deviation from the initial plan's literal Phase 3 snippet:</b> the plan's
/// Connect.cs sketch used <see cref="ClassInterfaceType.None"/>. That would
/// break this add-in's two actual COM entry paths: Word's RibbonX engine
/// resolves <c>onAction="OnPasteMarkdown"</c> by calling
/// <c>IDispatch.GetIDsOfNames("OnPasteMarkdown")</c> on this object at
/// runtime — it is late-bound by name, there is no other mechanism — and
/// <c>tools/e2e-test.ps1</c> calls
/// <c>$word.COMAddIns.Item(...).Object.PasteMarkdownFromClipboard()</c>, also
/// late-bound. Under <c>ClassInterfaceType.None</c> the COM-callable wrapper
/// only exposes IDispatch for the two explicitly-implemented interfaces
/// (<see cref="IDTExtensibility2"/>, <see cref="IRibbonExtensibility"/>) —
/// neither declares <c>OnPasteMarkdown</c>, <c>OnCopyMarkdown</c>, or
/// <c>PasteMarkdownFromClipboard</c>, so both the ribbon button and the e2e
/// script would fail to resolve the member ("member not found"), which is
/// the same class of "compiles clean, only dies at the live checkpoint"
/// failure the brief calls out for the assembly-resolution gap. Using
/// <see cref="ClassInterfaceType.AutoDispatch"/> instead generates a class
/// interface reflecting every public method, which both late-bound callers
/// need; Word's QueryInterface for the two explicit interface IIDs is
/// unaffected by this attribute either way. The versioning risk that makes
/// <c>None</c> the usual best practice (early-bound clients caching DISPIDs
/// across versions) does not apply here — every caller of this class
/// late-binds by name on every call. Unverified until the live checkpoint
/// (register-dev.ps1 → real Word); flagged in the Phase 3 report.
/// </para>
/// </remarks>
[ComVisible(true)]
[Guid("A18CE6CA-7818-468A-868D-69E34B4469F6")]
[ProgId("MdWord.AddIn.Connect")]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public class Connect : IDTExtensibility2, IRibbonExtensibility
{
    private dynamic _wordApp;

    /// <summary>
    /// A hidden control whose handle is created on Word's main UI thread
    /// during <see cref="OnConnection"/>, used to marshal work back onto
    /// that thread. Word's own <c>onAction</c> ribbon callbacks already run
    /// there, but <c>tools/e2e-test.ps1</c> reaches
    /// <see cref="PasteMarkdownFromClipboard"/> via
    /// <c>$word.COMAddIns.Item(...).Object</c>, which the COM/RPC layer
    /// dispatches on a separate MTA worker thread — confirmed empirically at
    /// the live checkpoint (thread-id/apartment-state logging showed
    /// <c>OnConnection</c> on the main STA thread but that external path on
    /// a different MTA thread). <see cref="System.Windows.Forms.Clipboard"/>
    /// requires an OLE-initialized STA thread; off the main thread it
    /// silently returns an empty string instead of throwing, so every paste
    /// driven that way saw an empty clipboard. <see cref="Control.Invoke"/>
    /// works without an owned message loop as long as the thread that
    /// created the handle is already pumping messages — true here, since
    /// that thread is Word's own UI thread.
    /// </summary>
    private Control _uiThreadMarshal;

    /// <summary>
    /// Installs the assembly-resolve handler before anything in this type is
    /// used. Word activates <see cref="Connect"/> via
    /// <c>mscoree.dll</c>/<c>CodeBase</c>, which only resolves
    /// <c>MdWord.AddIn.dll</c> itself — its dependencies
    /// (<c>MdWord.Core.dll</c>, Markdig, DocumentFormat.OpenXml.*, Jint,
    /// System.IO.Packaging) sit next to it in the same bin folder but are not
    /// found by normal .NET Framework probing (Word's install dir / the GAC).
    /// A static constructor runs before any member access on this type, so
    /// it is installed ahead of every code path — including
    /// <see cref="OnConnection"/> itself — that could eventually touch an
    /// <c>MdWord.Core</c> type via <see cref="WordActions"/>.
    /// </summary>
    static Connect()
    {
        AppDomain.CurrentDomain.AssemblyResolve += ResolveFromOwnDirectory;
    }

    private static Assembly ResolveFromOwnDirectory(object sender, ResolveEventArgs args)
    {
        try
        {
            var requestedName = new AssemblyName(args.Name).Name;
            var addInDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (addInDirectory == null)
            {
                return null;
            }

            var candidatePath = Path.Combine(addInDirectory, requestedName + ".dll");
            if (!File.Exists(candidatePath))
            {
                return null;
            }

            // All managed Word add-ins share the default AppDomain, and this
            // handler fires for every failed load of every add-in. Only serve
            // requests from our own dependency graph: the requesting assembly
            // lives in our folder, or there is no requester context at all
            // (mscoree/CodeBase activation of our own graph). Never satisfy
            // another add-in's failed load with our copy of a shared library.
            var requesterLocation = args.RequestingAssembly?.Location;
            if (!string.IsNullOrEmpty(requesterLocation) &&
                !string.Equals(Path.GetDirectoryName(requesterLocation), addInDirectory, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return Assembly.LoadFrom(candidatePath);
        }
        catch
        {
            // A resolve handler must never throw — return null (unresolved)
            // and let the normal FileNotFoundException surface instead.
            return null;
        }
    }

    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        try
        {
            _wordApp = Application;
            _uiThreadMarshal = new Control();
            var forceHandleCreation = _uiThreadMarshal.Handle;

            // e2e-test.ps1 hook: $word.COMAddIns.Item("MdWord.AddIn.Connect").Object
            // reaches this instance directly, letting the script call
            // PasteMarkdownFromClipboard() with no UI clicks involved.
            ((dynamic)AddInInst).Object = this;
        }
        catch (Exception ex)
        {
            HandleFailure("OnConnection", ex);
        }
    }

    public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
    {
        try
        {
            _wordApp = null;
            _uiThreadMarshal?.Dispose();
            _uiThreadMarshal = null;
        }
        catch (Exception ex)
        {
            HandleFailure("OnDisconnection", ex);
        }
    }

    public void OnAddInsUpdate(ref Array custom)
    {
    }

    public void OnStartupComplete(ref Array custom)
    {
        // Per the initial plan, Phase 6: pre-warm the KaTeX/Jint engine in the background so
        // the first formula paste doesn't pay the ~1s bundle-parse cost on the
        // UI thread. WordActions catches everything internally.
        System.Threading.Tasks.Task.Run((Action)WordActions.WarmUpMathEngine);
    }

    public void OnBeginShutdown(ref Array custom)
    {
    }

    public string GetCustomUI(string RibbonID)
    {
        try
        {
            return LoadEmbeddedRibbonXml();
        }
        catch (Exception ex)
        {
            HandleFailure("GetCustomUI", ex);
            return null;
        }
    }

    /// <summary>Ribbon-click entry point for the "Insert Markdown" button.</summary>
    public void OnPasteMarkdown(IRibbonControl control)
    {
        try
        {
            PasteMarkdownFromClipboard();
        }
        catch (Exception ex)
        {
            HandleFailure("OnPasteMarkdown", ex);
        }
    }

    /// <summary>Ribbon-click entry point for the "Copy as Markdown" button.</summary>
    public void OnCopyMarkdown(IRibbonControl control)
    {
        try
        {
            CopySelectionAsMarkdown();
        }
        catch (Exception ex)
        {
            HandleFailure("OnCopyMarkdown", ex);
        }
    }

    /// <summary>
    /// ComVisible orchestration entry point — called by both
    /// <see cref="OnPasteMarkdown"/> (ribbon click) and
    /// <c>tools/e2e-test.ps1</c> (<c>$addin.Object.PasteMarkdownFromClipboard()</c>
    /// via <c>AddInInst.Object</c>, set in <see cref="OnConnection"/>).
    /// </summary>
    public void PasteMarkdownFromClipboard()
    {
        try
        {
            if (_uiThreadMarshal != null && _uiThreadMarshal.InvokeRequired)
            {
                _uiThreadMarshal.Invoke(new Action(() => WordActions.PasteMarkdown(_wordApp)));
            }
            else
            {
                WordActions.PasteMarkdown(_wordApp);
            }
        }
        catch (Exception ex)
        {
            HandleFailure("PasteMarkdownFromClipboard", ex);
        }
    }

    /// <summary>
    /// ComVisible orchestration entry point — called by both
    /// <see cref="OnCopyMarkdown"/> (ribbon click) and
    /// <c>tools/e2e-test.ps1</c> (<c>$addin.Object.CopySelectionAsMarkdown()</c>
    /// via <c>AddInInst.Object</c>, set in <see cref="OnConnection"/>).
    /// </summary>
    public void CopySelectionAsMarkdown()
    {
        try
        {
            if (_uiThreadMarshal != null && _uiThreadMarshal.InvokeRequired)
            {
                _uiThreadMarshal.Invoke(new Action(() => WordActions.CopyMarkdown(_wordApp)));
            }
            else
            {
                WordActions.CopyMarkdown(_wordApp);
            }
        }
        catch (Exception ex)
        {
            HandleFailure("CopySelectionAsMarkdown", ex);
        }
    }

    private static string LoadEmbeddedRibbonXml()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = Array.Find(
            assembly.GetManifestResourceNames(),
            name => name.EndsWith("Ribbon.xml", StringComparison.Ordinal));

        if (resourceName == null)
        {
            throw new InvalidOperationException("Embedded resource 'Ribbon.xml' was not found in the MdWord.AddIn assembly.");
        }

        using (var stream = assembly.GetManifestResourceStream(resourceName))
        using (var reader = new StreamReader(stream))
        {
            return reader.ReadToEnd();
        }
    }

    /// <summary>
    /// The non-negotiable rule (initial plan, verbatim): every COM callback must
    /// log and show a human-readable MessageBox rather than let an exception
    /// escape — an escaped exception is what flips <c>LoadBehavior</c> from
    /// 3 to 2 and disables the add-in.
    /// </summary>
    private static void HandleFailure(string context, Exception exception)
    {
        Logger.LogError(context, exception);
        MessageBox.Show(
            $"MD-Word: an error occurred in '{context}'.\n{exception.Message}\n\nLog: {Logger.LogPath}",
            "MD-Word — error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
    }
}
