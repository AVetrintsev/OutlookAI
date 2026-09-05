using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace OutlookAI.Services.CustomActions
{
    public sealed class CustomActionStore
    {
        public const int CurrentSchemaVersion = 2;
        public const string TextEditingGroupId = "text_editing";
        private const string PersonalGroupId = "my_actions";

        public string Path { get; }

        public CustomActionStore()
            : this(ManifestPathResolver.AppDataFile("custom-actions.json"))
        {
        }

        public CustomActionStore(string path)
        {
            Path = path ?? throw new ArgumentNullException(nameof(path));
        }

        public CustomActionFile LoadCatalog()
        {
            EnsureFileExists();
            var file = JsonConvert.DeserializeObject<CustomActionFile>(File.ReadAllText(Path))
                ?? new CustomActionFile();
            var changed = false;
            if ((file.Groups == null || file.Groups.Length == 0) && file.Actions != null)
            {
                file = MigrateV1(file.Actions);
                changed = true;
            }

            file.SchemaVersion = CurrentSchemaVersion;
            var groups = (file.Groups ?? new CustomActionGroup[0])
                .Where(IsValid)
                .OrderBy(group => group.Order)
                .Select(group => group.Clone())
                .ToList();
            foreach (var baseline in ActionCatalog.Default.AssistantGroups().OrderBy(group => group.Order))
            {
                if (groups.Any(group =>
                    string.Equals(group.Id, baseline.Id, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                var insertAt = Math.Max(0, Math.Min(baseline.Order, groups.Count));
                groups.Insert(insertAt, baseline);
                changed = true;
            }

            foreach (var group in groups)
            {
                var originalActions = group.Actions ?? new CustomActionDefinition[0];
                var actions = originalActions.Where(IsValid).ToArray();
                if (actions.Length != originalActions.Length)
                {
                    changed = true;
                }
                if (actions.Any(action => NeedsNormalization(action, group.Id)))
                {
                    changed = true;
                }
                group.Actions = actions
                    .Select(action => NormalizeAction(action, group.Id))
                    .ToArray();
            }

            file.Groups = groups
                .Select((group, index) =>
                {
                    group.Order = index;
                    return group;
                })
                .ToArray();
            if (changed)
            {
                SaveCatalog(file);
            }
            return file;
        }

        public IReadOnlyList<CustomActionDefinition> Load()
        {
            return LoadCatalog().Groups
                .SelectMany(group => group.Actions ?? new CustomActionDefinition[0])
                .Where(action => !action.Disabled)
                .ToList();
        }

        public void SaveCatalog(CustomActionFile file)
        {
            if (file == null) throw new ArgumentNullException(nameof(file));
            var normalized = new CustomActionFile
            {
                SchemaVersion = CurrentSchemaVersion,
                Groups = (file.Groups ?? new CustomActionGroup[0])
                    .Where(IsValid)
                    .OrderBy(group => group.Order)
                    .Select((group, index) =>
                    {
                        var copy = group.Clone();
                        copy.Order = index;
                        copy.Actions = copy.Actions
                            .Where(IsValid)
                            .Select(action => NormalizeAction(action, copy.Id))
                            .ToArray();
                        return copy;
                    })
                    .ToArray()
            };
            var dir = System.IO.Path.GetDirectoryName(Path);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(Path, JsonConvert.SerializeObject(
                normalized,
                Formatting.Indented,
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore }));
        }

        public void Upsert(string groupId, CustomActionDefinition action)
        {
            if (!IsValid(action))
            {
                throw new ArgumentException("Action id, title, prompt and context are required.", nameof(action));
            }

            var file = LoadCatalog();
            var groups = file.Groups.ToList();
            var currentGroup = groups.FirstOrDefault(group =>
                group.Actions.Any(item => string.Equals(item.Id, action.Id, StringComparison.OrdinalIgnoreCase)));
            var target = groups.FirstOrDefault(group =>
                string.Equals(group.Id, groupId, StringComparison.OrdinalIgnoreCase));
            if (target == null)
            {
                target = groups.FirstOrDefault(group =>
                    string.Equals(group.Id, PersonalGroupId, StringComparison.OrdinalIgnoreCase));
            }
            if (target == null)
            {
                target = new CustomActionGroup
                {
                    Id = PersonalGroupId,
                    Title = "Мои действия",
                    Order = groups.Count,
                    Actions = new CustomActionDefinition[0]
                };
                groups.Add(target);
            }

            if (currentGroup != null && currentGroup != target)
            {
                currentGroup.Actions = currentGroup.Actions
                    .Where(item => !string.Equals(item.Id, action.Id, StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            }
            var actions = target.Actions.ToList();
            var index = actions.FindIndex(item =>
                string.Equals(item.Id, action.Id, StringComparison.OrdinalIgnoreCase));
            if (index >= 0) actions[index] = action;
            else actions.Add(action);
            target.Actions = actions.ToArray();
            file.Groups = groups.ToArray();
            SaveCatalog(file);
        }

        public void Upsert(CustomActionDefinition action)
        {
            Upsert(PersonalGroupId, action);
        }

        public bool Delete(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return false;
            var file = LoadCatalog();
            var removed = false;
            foreach (var group in file.Groups)
            {
                var remaining = group.Actions
                    .Where(action => !string.Equals(action.Id, id, StringComparison.OrdinalIgnoreCase))
                    .ToArray();
                removed |= remaining.Length != group.Actions.Length;
                group.Actions = remaining;
            }
            if (removed) SaveCatalog(file);
            return removed;
        }

        public bool Reorder(string groupId, IReadOnlyList<string> actionIds)
        {
            var file = LoadCatalog();
            var group = file.Groups.FirstOrDefault(item =>
                string.Equals(item.Id, groupId, StringComparison.OrdinalIgnoreCase));
            if (group == null) return false;
            var lookup = group.Actions.ToDictionary(item => item.Id, StringComparer.OrdinalIgnoreCase);
            var ordered = new List<CustomActionDefinition>();
            foreach (var id in actionIds ?? new string[0])
            {
                if (lookup.TryGetValue(id, out var action))
                {
                    ordered.Add(action);
                    lookup.Remove(id);
                }
            }
            ordered.AddRange(group.Actions.Where(action => lookup.ContainsKey(action.Id)));
            group.Actions = ordered.ToArray();
            SaveCatalog(file);
            return true;
        }

        public bool ResetGroup(string groupId)
        {
            var baseline = ActionCatalog.Default.AssistantGroups()
                .FirstOrDefault(group => string.Equals(group.Id, groupId, StringComparison.OrdinalIgnoreCase));
            if (baseline == null) return false;
            var file = LoadCatalog();
            var groups = file.Groups.ToList();
            var index = groups.FindIndex(group =>
                string.Equals(group.Id, groupId, StringComparison.OrdinalIgnoreCase));
            if (index >= 0) groups[index] = baseline;
            else groups.Add(baseline);
            file.Groups = groups.ToArray();
            SaveCatalog(file);
            return true;
        }

        public static string MakeActionId(string title)
        {
            var chars = (title ?? "").Trim().ToLowerInvariant()
                .Select(ch => char.IsLetterOrDigit(ch) ? ch : '_')
                .ToArray();
            var id = new string(chars).Trim('_');
            while (id.Contains("__")) id = id.Replace("__", "_");
            return string.IsNullOrWhiteSpace(id)
                ? "custom_action_" + DateTime.UtcNow.ToString("yyyyMMddHHmmss")
                : id;
        }

        private void EnsureFileExists()
        {
            if (File.Exists(Path)) return;
            SaveCatalog(new CustomActionFile
            {
                SchemaVersion = CurrentSchemaVersion,
                Groups = ActionCatalog.Default.AssistantGroups().ToArray()
            });
        }

        private static CustomActionFile MigrateV1(IEnumerable<CustomActionDefinition> actions)
        {
            var groups = ActionCatalog.Default.AssistantGroups().Select(group => group.Clone()).ToList();
            var migrated = (actions ?? Enumerable.Empty<CustomActionDefinition>())
                .Where(IsValid)
                .Select(action => action.Clone())
                .ToArray();
            if (migrated.Length > 0)
            {
                groups.Add(new CustomActionGroup
                {
                    Id = PersonalGroupId,
                    Title = "Мои действия",
                    Order = groups.Count,
                    Actions = migrated
                });
            }
            return new CustomActionFile
            {
                SchemaVersion = CurrentSchemaVersion,
                Groups = groups.ToArray()
            };
        }

        private static bool IsValid(CustomActionGroup group)
        {
            return group != null
                && !string.IsNullOrWhiteSpace(group.Id)
                && !string.IsNullOrWhiteSpace(group.Title)
                && group.Actions != null;
        }

        private static bool IsValid(CustomActionDefinition action)
        {
            return action != null
                && !string.IsNullOrWhiteSpace(action.Id)
                && !string.IsNullOrWhiteSpace(action.Title)
                && !string.IsNullOrWhiteSpace(action.Prompt)
                && action.Context != null;
        }

        private static bool NeedsNormalization(CustomActionDefinition action, string groupId)
        {
            var isEditorSelection = IsEditorSelectionAction(action, groupId);
            var expectedSurface = isEditorSelection
                ? CustomActionSurface.EditorSelection
                : CustomActionSurface.Assistant;
            if (!string.Equals(action.Surface, expectedSurface, StringComparison.Ordinal)) return true;
            if (!string.Equals(
                action.ApplicabilityItemType,
                CustomActionApplicability.NormalizeItemType(action.ApplicabilityItemType),
                StringComparison.Ordinal)) return true;
            if (!string.Equals(
                action.ApplicabilityDirection,
                CustomActionApplicability.NormalizeDirection(action.ApplicabilityDirection),
                StringComparison.Ordinal)) return true;
            if (!isEditorSelection) return false;
            return action.Context == null
                || !string.Equals(action.Context.Source, "selected_text", StringComparison.Ordinal)
                || action.Context.IncludeFullBodies
                || action.Context.IncludeAttachments
                || action.Context.MaxItems != 1
                || !string.Equals(action.Output, "replace_selection", StringComparison.Ordinal)
                || action.AllowTools
                || (action.AllowedTools?.Length ?? 0) != 0;
        }

        private static bool IsEditorSelectionAction(CustomActionDefinition action, string groupId)
        {
            if (string.Equals(groupId, TextEditingGroupId, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (string.Equals(
                CustomActionSurface.Normalize(action?.Surface),
                CustomActionSurface.EditorSelection,
                StringComparison.Ordinal))
            {
                return true;
            }
            return string.Equals(action?.Context?.Source, "selected_text", StringComparison.OrdinalIgnoreCase)
                && string.Equals(action?.Output, "replace_selection", StringComparison.OrdinalIgnoreCase);
        }

        private static CustomActionDefinition NormalizeAction(
            CustomActionDefinition action,
            string groupId)
        {
            var copy = action.Clone();
            var isEditorSelection = IsEditorSelectionAction(copy, groupId);
            copy.Surface = isEditorSelection
                ? CustomActionSurface.EditorSelection
                : CustomActionSurface.Assistant;
            copy.ApplicabilityItemType = CustomActionApplicability.NormalizeItemType(copy.ApplicabilityItemType);
            copy.ApplicabilityDirection = CustomActionApplicability.NormalizeDirection(copy.ApplicabilityDirection);
            if (isEditorSelection)
            {
                copy.Context.Source = "selected_text";
                copy.Context.MessageScope = "selected";
                copy.Context.IncludeFullBodies = false;
                copy.Context.IncludeAttachments = false;
                copy.Context.MaxItems = 1;
                copy.Output = "replace_selection";
                copy.AllowTools = false;
                copy.AllowedTools = new string[0];
            }
            return copy;
        }
    }
}
