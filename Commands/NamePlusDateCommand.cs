using EnvDTE;
using Microsoft.VisualStudio.Text;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace VSNamePlusDate
{
    [Command(PackageIds.NamePlusDateCommand)]
    internal sealed class NamePlusDateCommand : BaseCommand
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            string stamp = await CreateStampAsync();

            DocumentView? docView = await VS.Documents.GetActiveDocumentViewAsync();

            if (docView?.TextView != null)
            {
                SnapshotSpan? selection = docView.TextView.Selection.SelectedSpans.FirstOrDefault();

                if (selection.HasValue)
                {
                    docView.TextBuffer.Replace(selection.Value, stamp);
                }
                else
                {
                    int position = docView.TextView.Caret.Position.BufferPosition.Position;
                    docView.TextBuffer.Insert(position, stamp);
                }

                return;
            }

            Clipboard.SetText(stamp);
        }

        private async Task<string> CreateStampAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            DTE? dte = await VS.GetServiceAsync<DTE, DTE>();
            string? filePath = dte?.ActiveDocument?.FullName;
            string? solutionPath = dte?.Solution?.FullName;

            if (string.IsNullOrWhiteSpace(filePath))
                return $"VSNamePlusDate {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

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
                ".js" => $"// {relativePath}",
                ".ts" => $"// {relativePath}",
                ".json" => $"// {relativePath}",
                _ => $"// {relativePath}"
            };
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
            if (!path.EndsWith(Path.DirectorySeparatorChar.ToString()))
                return path + Path.DirectorySeparatorChar;

            return path;
        }
    }
}