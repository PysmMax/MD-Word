using System;
using System.IO;

namespace MdWord.AddIn;

/// <summary>
/// Minimal file logger to <c>%LOCALAPPDATA%\MD-Word\mdword.log</c> — no
/// external logging framework (PLAN.md Phase 3: "a handful of lines, not a
/// NuGet dependency"). Every COM callback in <see cref="Connect"/> logs here
/// before showing its <c>MessageBox</c>, so a disabled add-in
/// (<c>LoadBehavior</c> flipping from 3 to 2 after a thrown
/// <c>OnConnection</c>) is always diagnosable from this file.
/// </summary>
internal static class Logger
{
    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MD-Word");

    private static readonly string LogFilePath = Path.Combine(LogDirectory, "mdword.log");

    private static readonly object WriteLock = new object();

    /// <summary>Absolute path to the log file — surfaced in error MessageBoxes.</summary>
    public static string LogPath => LogFilePath;

    public static void LogError(string context, Exception exception)
    {
        Write($"ERROR [{context}]: {exception}");
    }

    public static void LogWarning(string message)
    {
        Write($"WARN: {message}");
    }

    private static void Write(string line)
    {
        try
        {
            lock (WriteLock)
            {
                Directory.CreateDirectory(LogDirectory);
                RotateIfTooBig();
                File.AppendAllText(LogFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {line}{Environment.NewLine}");
            }
        }
        catch
        {
            // Logging must never itself throw and mask the original failure —
            // a broken log path is not a reason to crash a COM callback.
        }
    }

    /// <summary>
    /// Caps the log at ~1 MB by rolling it to a single ".old" generation —
    /// enough history for diagnostics without unbounded growth.
    /// </summary>
    private static void RotateIfTooBig()
    {
        const long MaxLogBytes = 1024 * 1024;
        var info = new FileInfo(LogFilePath);
        if (!info.Exists || info.Length < MaxLogBytes)
        {
            return;
        }

        var backupPath = LogFilePath + ".old";
        File.Delete(backupPath); // no-op when the file does not exist
        File.Move(LogFilePath, backupPath);
    }
}
