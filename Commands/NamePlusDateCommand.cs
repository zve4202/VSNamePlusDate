using EnvDTE;
using Microsoft.VisualStudio.Text;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace VSNamePlusDate
{
    [Command(PackageIds.NamePlusDateCommand)]
    internal sealed class NamePlusDateCommand : BaseCommand<NamePlusDateCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            DocumentView? docView = await VS.Documents.GetActiveDocumentViewAsync();

            SnapshotSpan? selection = docView?.TextView.Selection.SelectedSpans.FirstOrDefault();            
            
            if (docView == null || selection == null)
                return;

            string name = await GetSolutionNameAsync();

            if (string.IsNullOrWhiteSpace(name))
                name = "Project";

            string stamp = $"{name} {DateTime.Now:yyyy-MM-dd HH:mm:ss}";

            docView.TextBuffer.Replace(selection.Value, stamp);
        }

        private async Task<string> GetSolutionNameAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            DTE? dte = await VS.GetServiceAsync<DTE, DTE>();

            string? solutionPath = dte?.Solution?.FullName;

            if (string.IsNullOrWhiteSpace(solutionPath))
                return string.Empty;

            return Path.GetFileNameWithoutExtension(solutionPath);
        }
    }
}
