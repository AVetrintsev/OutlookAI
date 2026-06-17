using System.Linq;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.CustomActions
{
    public static class CustomActionUiSerializer
    {
        public static JObject Build(CustomActionFile catalog, string[] recommendationIds, JArray tools)
        {
            var resettableGroups = new System.Collections.Generic.HashSet<string>(
                ActionCatalog.Default.AssistantGroups().Select(group => group.Id),
                System.StringComparer.OrdinalIgnoreCase);
            var groups = new JArray((catalog?.Groups ?? new CustomActionGroup[0])
                .OrderBy(group => group.Order)
                .Select(group => new JObject(
                    new JProperty("id", group.Id),
                    new JProperty("title", group.Title),
                    new JProperty("can_reset", resettableGroups.Contains(group.Id)),
                    new JProperty("actions", new JArray((group.Actions ?? new CustomActionDefinition[0])
                        .Where(action => !action.Disabled)
                        .Select(ToJson))))));
            return new JObject(
                new JProperty("groups", groups),
                new JProperty("recommendations_enabled", Config.RecommendationsEnabled),
                new JProperty("recommendation_ids", new JArray(recommendationIds ?? new string[0])),
                new JProperty("tools", tools ?? new JArray()));
        }

        public static JObject ToJson(CustomActionDefinition action)
        {
            var output = action.Output == "create_draft" ? "create_reply" : action.Output ?? "chat";
            return new JObject(
                new JProperty("id", action.Id),
                new JProperty("label", action.Title),
                new JProperty("description", action.Description ?? ""),
                new JProperty("action_prompt", action.Prompt ?? ""),
                new JProperty("source", action.Context?.Source ?? "current_selection"),
                new JProperty("read_filter", action.Context?.ReadFilter ?? "all"),
                new JProperty("time_range", action.Context?.TimeRange ?? "today"),
                new JProperty("manual_from", action.Context?.ManualFrom?.ToString("o")),
                new JProperty("manual_to", action.Context?.ManualTo?.ToString("o")),
                new JProperty("max_items", action.Context?.MaxItems ?? 20),
                new JProperty("include_full_bodies", action.Context?.IncludeFullBodies ?? true),
                new JProperty("include_attachments", action.Context?.IncludeAttachments ?? false),
                new JProperty("output", output),
                new JProperty("allowed_tools", new JArray(action.AllowedTools ?? new string[0])));
        }
    }
}
