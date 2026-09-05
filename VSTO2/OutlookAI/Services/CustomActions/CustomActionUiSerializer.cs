using System.Linq;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.CustomActions
{
    public static class CustomActionUiSerializer
    {
        public static JObject Build(
            CustomActionFile catalog,
            string[] recommendationIds,
            JArray tools,
            CustomActionApplicabilityContext applicability = null,
            bool includeEditorSelectionActions = true)
        {
            var emptyTextEditingGroup = includeEditorSelectionActions
                ? (catalog?.Groups ?? new CustomActionGroup[0])
                    .FirstOrDefault(group => string.Equals(
                        group.Id,
                        CustomActionStore.TextEditingGroupId,
                        System.StringComparison.OrdinalIgnoreCase))
                    ?.Clone()
                : null;
            catalog = CustomActionApplicability.FilterCatalog(catalog, applicability);
            var resettableGroups = new System.Collections.Generic.HashSet<string>(
                ActionCatalog.Default.AssistantGroups().Select(group => group.Id),
                System.StringComparer.OrdinalIgnoreCase);
            var visibleGroups = (catalog?.Groups ?? new CustomActionGroup[0])
                .OrderBy(group => group.Order)
                .Select(group =>
                {
                    var copy = group.Clone();
                    if (!includeEditorSelectionActions)
                    {
                        copy.Actions = copy.Actions
                            .Where(action => CustomActionSurface.Normalize(action.Surface)
                                != CustomActionSurface.EditorSelection)
                            .ToArray();
                    }
                    return copy;
                })
                .Where(group => group.Actions.Length > 0)
                .ToList();
            if (emptyTextEditingGroup != null
                && !visibleGroups.Any(group => string.Equals(
                    group.Id,
                    CustomActionStore.TextEditingGroupId,
                    System.StringComparison.OrdinalIgnoreCase)))
            {
                emptyTextEditingGroup.Actions = new CustomActionDefinition[0];
                visibleGroups.Add(emptyTextEditingGroup);
                visibleGroups = visibleGroups.OrderBy(group => group.Order).ToList();
            }
            var recommendationActionIds = new System.Collections.Generic.HashSet<string>(
                visibleGroups
                    .SelectMany(group => group.Actions)
                    .Where(action => CustomActionSurface.Normalize(action.Surface)
                        == CustomActionSurface.Assistant)
                    .Select(action => action.Id),
                System.StringComparer.OrdinalIgnoreCase);
            var groups = new JArray(visibleGroups
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
                new JProperty("recommendation_ids", new JArray((recommendationIds ?? new string[0])
                    .Where(recommendationActionIds.Contains))),
                new JProperty("tools", tools ?? new JArray()));
        }

        public static JObject ToJson(CustomActionDefinition action)
        {
            var output = action.Output == "create_draft" ? "create_reply" : action.Output ?? "chat";
            var itemType = CustomActionApplicability.NormalizeItemType(action.ApplicabilityItemType);
            var direction = CustomActionApplicability.NormalizeDirection(action.ApplicabilityDirection);
            return new JObject(
                new JProperty("id", action.Id),
                new JProperty("label", action.Title),
                new JProperty("description", action.Description ?? ""),
                new JProperty("action_prompt", action.Prompt ?? ""),
                new JProperty("surface", CustomActionSurface.Normalize(action.Surface)),
                new JProperty("use_skills", action.UseSkills),
                new JProperty("source", action.Context?.Source ?? "current_selection"),
                new JProperty("read_filter", action.Context?.ReadFilter ?? "all"),
                new JProperty("time_range", action.Context?.TimeRange ?? "today"),
                new JProperty("manual_from", action.Context?.ManualFrom?.ToString("o")),
                new JProperty("manual_to", action.Context?.ManualTo?.ToString("o")),
                new JProperty("max_items", action.Context?.MaxItems ?? 20),
                new JProperty("include_full_bodies", action.Context?.IncludeFullBodies ?? true),
                new JProperty("include_attachments", action.Context?.IncludeAttachments ?? false),
                new JProperty("output", output),
                new JProperty("applicability_item_type", itemType),
                new JProperty("applicability_direction", direction),
                new JProperty("allowed_tools", new JArray(action.AllowedTools ?? new string[0])));
        }
    }
}
