// Commands\CreateCommentCommand.cs
using Community.VisualStudio.Toolkit;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using WindowEvents = EnvDTE.WindowEvents;


namespace VSExtension;

[Command(PackageIds.CreateCommentCommand)]
internal sealed class CreateCommentCommand : BaseCommand<CreateCommentCommand>
{
    private static WindowEvents? _windowEvents;
    private static bool _isAutoRunning;
    private static DateTime _lastAutoRun = DateTime.MinValue;

    protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
    {
        await RunAsync();
    }

    internal static async Task StartGitChangesAutoPasteAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        DTE? dte = await VS.GetServiceAsync<DTE, DTE>();

        if (dte == null)
            return;

        _windowEvents = dte.Events.WindowEvents;

        _windowEvents.WindowActivated += async (gotFocus, lostFocus) =>
        {
            await OnWindowActivatedAsync(gotFocus);
        };
    }

    private static async Task OnWindowActivatedAsync(Window gotFocus)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        if (_isAutoRunning)
            return;

        if ((DateTime.Now - _lastAutoRun).TotalSeconds < 2)
            return;

        string caption = gotFocus.Caption ?? string.Empty;

        bool isGitChanges =
            caption.Contains("Git Changes", StringComparison.OrdinalIgnoreCase) ||
            caption.Contains("Изменения Git", StringComparison.OrdinalIgnoreCase);

        if (!isGitChanges)
            return;

        _isAutoRunning = true;
        _lastAutoRun = DateTime.Now;

        try
        {
            await Task.Delay(300);
            await RunAsync(forcePaste: true);
        }
        finally
        {
            _isAutoRunning = false;
        }
    }

    private static async Task RunAsync(bool forcePaste = false)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        string stamp = await CreateStampAsync();

        if (!forcePaste && IsFocusInCodeEditor())
        {
            DocumentView? docView = await VS.Documents.GetActiveDocumentViewAsync();

            if (docView?.TextView != null && docView.TextBuffer != null)
            {
                InsertIntoEditor(docView, stamp);
                return;
            }
        }

        await PasteIntoFocusedControlAsync(stamp);
    }

    private static void InsertIntoEditor(DocumentView docView, string stamp)
    {
        IWpfTextView textView = docView.TextView;
        ITextBuffer textBuffer = docView.TextBuffer;

        string textToInsert = stamp + Environment.NewLine;

        textBuffer.Insert(0, textToInsert);

        textView.Caret.MoveTo(
            new SnapshotPoint(textBuffer.CurrentSnapshot, textToInsert.Length)
        );
    }

    private static async Task PasteIntoFocusedControlAsync(string stamp)
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        Clipboard.SetText(stamp);

        await Task.Delay(100);

        SendKeys.SendWait("^v");
    }

    private static bool IsFocusInCodeEditor()
    {
        IntPtr hwnd = GetFocus();

        if (hwnd == IntPtr.Zero)
            return false;

        string className = GetWindowClassName(hwnd);

        return className.Contains("WpfTextView", StringComparison.OrdinalIgnoreCase)
            || className.Contains("VsTextEditPane", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetWindowClassName(IntPtr hwnd)
    {
        const int maxCount = 256;

        StringBuilder className = new(maxCount);

        GetClassName(hwnd, className, maxCount);

        return className.ToString();
    }

    private static async Task<string> CreateStampAsync()
    {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        DTE? dte = await VS.GetServiceAsync<DTE, DTE>();

        string? filePath = dte?.ActiveDocument?.FullName;
        string? solutionPath = dte?.Solution?.FullName;

        if (string.IsNullOrWhiteSpace(filePath))
        {
            string solutionName = GetSolutionName(solutionPath);
            return $"{solutionName} {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        }

        string relativePath = GetRelativePath(solutionPath, filePath);
        relativePath = relativePath.Replace("/", "\\");

        string extension = Path.GetExtension(filePath).ToLowerInvariant();

        return extension switch
        {
            ".razor" => $"@* {relativePath} *@",
            ".cshtml" => $"@* {relativePath} *@",

            ".html" => $"<!-- {relativePath} -->",
            ".xml" => $"<!-- {relativePath} -->",
            ".xaml" => $"<!-- {relativePath} -->",

            ".css" => $"/* {relativePath} */",
            ".scss" => $"/* {relativePath} */",
            ".less" => $"/* {relativePath} */",

            ".sql" => $"-- {relativePath}",

            ".vb" => $"' {relativePath}",
            ".ps1" => $"# {relativePath}",

            ".cs" => $"// {relativePath}",
            ".js" => $"// {relativePath}",
            ".ts" => $"// {relativePath}",
            ".jsx" => $"// {relativePath}",
            ".tsx" => $"// {relativePath}",
            ".json" => $"// {relativePath}",

            _ => $"// {relativePath}"
        };
    }

    private static string GetSolutionName(string? solutionPath)
    {
        if (string.IsNullOrWhiteSpace(solutionPath))
            return "Project";

        string? name = Path.GetFileNameWithoutExtension(solutionPath);

        return string.IsNullOrWhiteSpace(name) ? "Project" : name;
    }

    private static string GetRelativePath(string? solutionPath, string filePath)
    {
        if (string.IsNullOrWhiteSpace(solutionPath))
            return Path.GetFileName(filePath);

        string? solutionDir = Path.GetDirectoryName(solutionPath);

        if (string.IsNullOrWhiteSpace(solutionDir))
            return Path.GetFileName(filePath);

        Uri baseUri = new Uri(AppendDirectorySeparatorChar(solutionDir));
        Uri fileUri = new Uri(filePath);

        return Uri.UnescapeDataString(
            baseUri.MakeRelativeUri(fileUri).ToString()
        );
    }

    private static string AppendDirectorySeparatorChar(string path)
    {
        return path.EndsWith(Path.DirectorySeparatorChar.ToString())
            ? path
            : path + Path.DirectorySeparatorChar;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetFocus();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
}