using System;
using System.Linq;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services.Assistant
{
    public sealed class AssistantPlanValidator : IPlanValidator
    {
        private readonly WorkflowManifestCatalog _workflows;

        public AssistantPlanValidator()
            : this(WorkflowManifestCatalog.Default)
        {
        }

        public AssistantPlanValidator(WorkflowManifestCatalog workflows)
        {
            _workflows = workflows ?? WorkflowManifestCatalog.Default;
        }

        public ValidatedTurnPlan Validate(
            TurnPlan plan,
            TurnRequest request,
            TurnSnapshot snapshot,
            bool includeWriteTools)
        {
            plan = plan ?? new TurnPlan();
            var workflow = _workflows.Get(plan.Workflow ?? request?.WorkflowId);
            var allowedGroups = new System.Collections.Generic.HashSet<string>(
                workflow.AllowedCapabilityGroups ?? new string[0],
                StringComparer.OrdinalIgnoreCase);
            var selected = (snapshot?.Capabilities ?? new CapabilitySearchResult[0])
                .Where(result => result.Card != null)
                .Where(result => allowedGroups.Count == 0 || allowedGroups.Contains(result.Card.Group ?? ""))
                .Where(result => plan.CapabilityIds == null
                    || plan.CapabilityIds.Length == 0
                    || plan.CapabilityIds.Contains(result.Card.Id, StringComparer.OrdinalIgnoreCase))
                .ToArray();

            if (plan.NeedsMoreData)
            {
                return Clarification(plan, plan.MissingData);
            }

            if (NeedsSelectedItem(plan) && !HasSelection(snapshot))
            {
                return Clarification(plan, "Выберите письмо, о котором идет речь.");
            }

            var contextMode = NormalizeContextMode(plan.ContextMode);
            if (contextMode == "metadata" && LooksContentHeavy(plan))
            {
                contextMode = "snippet";
            }
            if (LooksSearchOrReport(plan))
            {
                contextMode = "search_projection";
            }

            var tools = selected
                .SelectMany(result => result.Card.ToolNames ?? new string[0])
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Where(name => includeWriteTools || !ToolManifestCatalog.IsWriteTool(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (tools.Length == 0 && contextMode == "snippet")
            {
                tools = new[] { "outlook_get_current_selection" };
            }
            if (contextMode == "snippet" && HasSelection(snapshot))
                tools = tools.Concat(new[] { "outlook_read_conversation" })
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            if (LooksDraftIntent(plan))
            {
                contextMode = "snippet";
                tools = tools
                    .Concat(new[] { "outlook_create_draft" })
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            }
            if (tools.Length == 0 && contextMode == "search_projection")
            {
                tools = new[] { "outlook_search_messages", "outlook_count_messages", "outlook_aggregate_messages", "outlook_list_folders" };
            }
            if (contextMode == "metadata")
            {
                tools = new string[0];
            }
            return new ValidatedTurnPlan
            {
                Plan = plan,
                ContextMode = contextMode,
                AllowedToolNames = tools,
                IncludeWriteTools = includeWriteTools && tools.Any(ToolManifestCatalog.IsWriteTool),
                SelectedCapabilities = selected
            };
        }

        private static ValidatedTurnPlan Clarification(TurnPlan plan, string missingData)
        {
            return new ValidatedTurnPlan
            {
                Plan = plan,
                ContextMode = "metadata",
                AllowedToolNames = new string[0],
                NeedsMoreData = true,
                MissingData = string.IsNullOrWhiteSpace(missingData)
                    ? "Мне не хватает данных, чтобы ответить уверенно."
                    : missingData
            };
        }

        private static bool HasSelection(TurnSnapshot snapshot)
        {
            return snapshot?.Selection != null
                && snapshot.Selection.Count > 0
                && snapshot.Selection.Messages != null
                && snapshot.Selection.Messages.Count > 0;
        }

        private static bool NeedsSelectedItem(TurnPlan plan)
        {
            var intent = (plan.Intent ?? "").ToLowerInvariant();
            return intent.Contains("selected") || intent.Contains("message") || intent.Contains("content");
        }

        private static bool LooksContentHeavy(TurnPlan plan)
        {
            var text = ((plan.Intent ?? "") + " " + (plan.Reason ?? "") + " " + (plan.OutputContract ?? "")).ToLowerInvariant();
            return text.Contains("content")
                || text.Contains("summary")
                || text.Contains("summar")
                || text.Contains("reply")
                || text.Contains("draft")
                || text.Contains("why")
                || text.Contains("purpose");
        }

        private static bool LooksSearchOrReport(TurnPlan plan)
        {
            var text = ((plan.Intent ?? "") + " " + (plan.OutputContract ?? "")).ToLowerInvariant();
            return text.Contains("search")
                || text.Contains("report")
                || text.Contains("export")
                || text.Contains("find");
        }

        private static bool LooksDraftIntent(TurnPlan plan)
        {
            var text = ((plan.Intent ?? "") + " " + (plan.OutputContract ?? "")).ToLowerInvariant();
            return text.Contains("create_draft")
                || text.Contains("draft")
                || text.Contains("reply_draft");
        }

        private static string NormalizeContextMode(string value)
        {
            value = (value ?? "").Trim().ToLowerInvariant();
            return value == "snippet" || value == "full" || value == "search_projection"
                ? value
                : "metadata";
        }
    }
}
