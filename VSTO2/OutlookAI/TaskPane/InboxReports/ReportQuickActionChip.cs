using System.Collections.Generic;
using System.Linq;
using OutlookAI.Services;

namespace OutlookAI.TaskPane.InboxReports
{
    /// <summary>
    /// One quick-action chip in the Inbox Reports pane. Definitions live in
    /// Prompts/builtin-actions.json.
    /// </summary>
    public sealed class ReportQuickActionChip
    {
        public string Label { get; set; }
        public string TemplateText { get; set; }

        public static IReadOnlyList<ReportQuickActionChip> Defaults()
        {
            return ActionCatalog.Default.ForSurface("inbox_reports")
                .Select(a => new ReportQuickActionChip { Label = a.Label, TemplateText = a.Prompt })
                .ToList();
        }
    }
}
