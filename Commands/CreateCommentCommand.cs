// Commands\CreateCommentCommand.csusing Community.VisualStudio.Toolkit;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace VSExtension
{
    [Command(PackageIds.CreateCommentCommand)]
    internal sealed class CreateCommentCommand : BaseCommand<CreateCommentCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            string stamp = await CreateStampAsync();

            DocumentView? docView = await VS.Documents.GetActiveDocumentViewAsync();

            if (docView?.TextView != null && docView?.TextBuffer != null)
            {
                InsertIntoEditor(docView, stamp);
                return;
            }

            CopyToClipboardAndPasteIntoGitChanges(stamp);
        }

        private static void InsertIntoEditor(DocumentView docView, string stamp)
        {
            IWpfTextView textView = docView.TextView;
            ITextBuffer textBuffer = docView.TextBuffer;

            string textToInsert =
                stamp +
                Environment.NewLine +
                Environment.NewLine;

            textBuffer.Insert(0, textToInsert);

            textView.Caret.MoveTo(
                new SnapshotPoint(
                    textBuffer.CurrentSnapshot,
                    textToInsert.Length
                )
            );
        }
        private static async void CopyToClipboardAndPasteIntoGitChanges(string stamp)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            Clipboard.SetText(stamp);

            await VS.Commands.ExecuteAsync("View.GitChanges");

            await Task.Delay(750);

            SendKeys.SendWait("^v");
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

                ".vb" => $"'{relativePath}",
                ".ps1" => $"# {relativePath}",

                ".js" => $"// {relativePath}",
                ".ts" => $"// {relativePath}",
                ".jsx" => $"// {relativePath}",
                ".tsx" => $"// {relativePath}",
                ".json" => $"// {relativePath}",
                ".cs" => $"// {relativePath}",

                _ => $"// {relativePath}"
            };
        }

        private static string GetSolutionName(string? solutionPath)
        {
            if (string.IsNullOrWhiteSpace(solutionPath))
                return "Project";

            string? name = Path.GetFileNameWithoutExtension(solutionPath);

            if (string.IsNullOrWhiteSpace(name))
                return "Project";

            return name;
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
            if (path.EndsWith(Path.DirectorySeparatorChar.ToString()))
                return path;

            return path + Path.DirectorySeparatorChar;
        }
    }
}