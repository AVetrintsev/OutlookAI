using System.Collections.Generic;
using System.Linq;
using OutlookAI.Services;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// A pre-canned prompt rendered as a clickable chip above the Inbox
    /// Copilot composer. Definitions live in Prompts/builtin-actions.json.
    /// </summary>
    public sealed class QuickActionChip
    {
        public string Label { get; set; }
        public string Prompt { get; set; }

        public static IReadOnlyList<QuickActionChip> ComputeChipsForSelectionCount(int selectionCount)
        {
            return ActionCatalog.Default.ForSurface("inbox_copilot")
                .Where(a => a.Selection == "any"
                    || (selectionCount == 1 && a.Selection == "single")
                    || (selectionCount >= 2 && a.Selection == "multiple"))
                .Select(a => new QuickActionChip { Label = a.Label, Prompt = a.Prompt })
                .ToList();
        }
    }
}
