using System;
using System.Runtime.InteropServices;

namespace MdWord.AddIn;

/// <summary>
/// Manual <c>ComImport</c> declarations for the three stable Office ABI
/// interfaces this add-in needs (shared-add-in extensibility + RibbonX),
/// instead of referencing the <c>Extensibility</c>/<c>Microsoft.Office.Core</c>
/// PIAs — three fixed GUIDs, zero extra assembly dependencies (PLAN.md Phase 3
/// "Технології"). These GUIDs are the well-known, stable Office values (not
/// re-derivable from the canonical MS Learn "customize the ribbon by using a
/// managed COM add-in" walkthrough fetched for this phase's Step 0 — that
/// article uses the typed <c>Extensibility</c> PIA reference rather than raw
/// <c>ComImport</c>, so it never prints the literal GUID strings; it does
/// confirm the same interface shapes and calling pattern, e.g. storing the
/// <c>Application</c> object in <c>OnConnection</c> and the
/// <c>GetCustomUI</c>/embedded-resource idiom used in <see cref="Connect"/>).
/// </summary>
[ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"),
 InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IDTExtensibility2
{
    [DispId(1)]
    void OnConnection(
        [In, MarshalAs(UnmanagedType.IDispatch)] object Application,
        [In] ext_ConnectMode ConnectMode,
        [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
        [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

    [DispId(2)]
    void OnDisconnection(
        [In] ext_DisconnectMode RemoveMode,
        [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

    [DispId(3)]
    void OnAddInsUpdate(
        [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

    [DispId(4)]
    void OnStartupComplete(
        [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

    [DispId(5)]
    void OnBeginShutdown(
        [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
}

public enum ext_ConnectMode
{
    ext_cm_AfterStartup = 0,
    ext_cm_Startup = 1,
    ext_cm_External = 2,
    ext_cm_CommandLine = 3,
}

public enum ext_DisconnectMode
{
    ext_dm_HostShutdown = 0,
    ext_dm_UserClosed = 1,
}

[ComImport, Guid("000C0396-0000-0000-C000-000000000046"),
 InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRibbonExtensibility
{
    [DispId(1)]
    string GetCustomUI(string RibbonID);
}

[ComImport, Guid("000C0395-0000-0000-C000-000000000046"),
 InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRibbonControl
{
    [DispId(1)]
    string Id { get; }

    [DispId(2)]
    object Context { get; }

    [DispId(3)]
    string Tag { get; }
}
