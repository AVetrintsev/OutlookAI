using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace OutlookAI.Services.CustomActions
{
    public sealed class CustomActionImportPreview
    {
        public CustomActionFile Imported { get; set; }
        public string[] Replacements { get; set; }
    }

    public static class CustomActionYamlCodec
    {
        public static string Export(CustomActionFile file, string actionId = null)
        {
            var sb = new StringBuilder();
            sb.AppendLine("schema_version: 2");
            sb.AppendLine("groups:");
            foreach (var group in (file?.Groups ?? new CustomActionGroup[0]).OrderBy(item => item.Order))
            {
                var actions = group.Actions ?? new CustomActionDefinition[0];
                if (!string.IsNullOrWhiteSpace(actionId))
                {
                    actions = actions.Where(action =>
                        string.Equals(action.Id, actionId, StringComparison.OrdinalIgnoreCase)).ToArray();
                    if (actions.Length == 0) continue;
                }
                sb.AppendLine("  - id: " + Scalar(group.Id));
                sb.AppendLine("    title: " + Scalar(group.Title));
                sb.AppendLine("    actions:");
                foreach (var action in actions)
                {
                    sb.AppendLine("      - id: " + Scalar(action.Id));
                    sb.AppendLine("        title: " + Scalar(action.Title));
                    sb.AppendLine("        description: " + Scalar(action.Description ?? ""));
                    sb.AppendLine("        prompt: |");
                    foreach (var line in (action.Prompt ?? "").Replace("\r", "").Split('\n'))
                    {
                        sb.AppendLine("          " + line);
                    }
                    sb.AppendLine("        surface: " +
                        Scalar(CustomActionSurface.Normalize(action.Surface)));
                    sb.AppendLine("        output: " + Scalar(action.Output ?? "chat"));
                    sb.AppendLine("        use_skills: " + (action.UseSkills ? "true" : "false"));
                    sb.AppendLine("        applicability_item_type: " +
                        Scalar(CustomActionApplicability.NormalizeItemType(action.ApplicabilityItemType)));
                    sb.AppendLine("        applicability_direction: " +
                        Scalar(CustomActionApplicability.NormalizeDirection(action.ApplicabilityDirection)));
                    sb.AppendLine("        disabled: " + (action.Disabled ? "true" : "false"));
                    sb.AppendLine("        context_json: " + Scalar(JsonConvert.SerializeObject(action.Context)));
                    sb.AppendLine("        allowed_tools_json: " +
                        Scalar(JsonConvert.SerializeObject(action.AllowedTools ?? new string[0])));
                }
            }
            return sb.ToString();
        }

        public static CustomActionImportPreview Preview(string yaml, CustomActionFile current)
        {
            var imported = Parse(yaml);
            var existing = new HashSet<string>(
                (current?.Groups ?? new CustomActionGroup[0])
                    .SelectMany(group => (group.Actions ?? new CustomActionDefinition[0])
                        .Select(action => MatchKey(group.Title, action.Title))),
                StringComparer.OrdinalIgnoreCase);
            var replacements = imported.Groups
                .SelectMany(group => group.Actions
                    .Where(action => existing.Contains(MatchKey(group.Title, action.Title)))
                    .Select(action => group.Title + " / " + action.Title))
                .ToArray();
            return new CustomActionImportPreview
            {
                Imported = imported,
                Replacements = replacements
            };
        }

        public static CustomActionFile Merge(CustomActionFile current, CustomActionFile imported)
        {
            var groups = (current?.Groups ?? new CustomActionGroup[0])
                .Select(group => group.Clone())
                .ToList();
            foreach (var incomingGroup in imported?.Groups ?? new CustomActionGroup[0])
            {
                var target = groups.FirstOrDefault(group =>
                    Normalize(group.Title) == Normalize(incomingGroup.Title));
                if (target == null)
                {
                    var copy = incomingGroup.Clone();
                    copy.Order = groups.Count;
                    groups.Add(copy);
                    continue;
                }

                var actions = target.Actions.ToList();
                foreach (var incoming in incomingGroup.Actions)
                {
                    var index = actions.FindIndex(action =>
                        Normalize(action.Title) == Normalize(incoming.Title));
                    if (index >= 0)
                    {
                        incoming.Id = actions[index].Id;
                        actions[index] = incoming.Clone();
                    }
                    else
                    {
                        if (actions.Any(action => string.Equals(action.Id, incoming.Id, StringComparison.OrdinalIgnoreCase)))
                        {
                            incoming.Id = CustomActionStore.MakeActionId(incoming.Title + "_" + DateTime.UtcNow.Ticks);
                        }
                        actions.Add(incoming.Clone());
                    }
                }
                target.Actions = actions.ToArray();
            }
            return new CustomActionFile
            {
                SchemaVersion = CustomActionStore.CurrentSchemaVersion,
                Groups = groups.ToArray()
            };
        }

        public static CustomActionFile Parse(string yaml)
        {
            if (string.IsNullOrWhiteSpace(yaml)) throw new FormatException("YAML is empty.");
            var lines = yaml.Replace("\r", "").Split('\n');
            if (!lines.Any(line => string.Equals(line.Trim(), "groups:", StringComparison.OrdinalIgnoreCase)))
            {
                var firstGroup = Array.FindIndex(lines, line => line.StartsWith("- title: "));
                if (firstGroup >= 0)
                {
                    lines = lines.Skip(firstGroup).Select(line =>
                    {
                        if (line.StartsWith("- title: ")) return "  " + line;
                        if (line.StartsWith("  items:")) return "    actions:";
                        if (line.StartsWith("    ")) return "  " + line;
                        return line;
                    }).ToArray();
                }
            }
            var groups = new List<CustomActionGroup>();
            CustomActionGroup group = null;
            CustomActionDefinition action = null;
            var prompt = new StringBuilder();
            var readingPrompt = false;

            Action flushPrompt = () =>
            {
                if (action != null && readingPrompt)
                {
                    action.Prompt = prompt.ToString().TrimEnd();
                }
                prompt.Clear();
                readingPrompt = false;
            };

            foreach (var raw in lines)
            {
                var line = raw ?? "";
                if (readingPrompt)
                {
                    if (line.StartsWith("          "))
                    {
                        if (prompt.Length > 0) prompt.AppendLine();
                        prompt.Append(line.Substring(10));
                        continue;
                    }
                    flushPrompt();
                }

                var trimmed = line.Trim();
                if (trimmed.Length == 0 || trimmed.StartsWith("#")) continue;
                var indent = line.Length - line.TrimStart().Length;

                if (indent == 2 && trimmed.StartsWith("- "))
                {
                    group = new CustomActionGroup
                    {
                        Id = "",
                        Title = "",
                        Order = groups.Count,
                        Actions = new CustomActionDefinition[0]
                    };
                    groups.Add(group);
                    action = null;
                    AssignGroup(group, trimmed.Substring(2));
                    continue;
                }
                if (indent == 4 && group != null)
                {
                    AssignGroup(group, trimmed);
                    continue;
                }
                if (indent == 6 && trimmed.StartsWith("- ") && group != null)
                {
                    action = DefaultAction();
                    group.Actions = group.Actions.Concat(new[] { action }).ToArray();
                    AssignAction(action, trimmed.Substring(2));
                    continue;
                }
                if (indent == 8 && action != null)
                {
                    if (trimmed == "prompt: |" || trimmed == "prompt: >")
                    {
                        readingPrompt = true;
                        continue;
                    }
                    AssignAction(action, trimmed);
                }
            }
            flushPrompt();

            foreach (var item in groups)
            {
                if (string.IsNullOrWhiteSpace(item.Title)) throw new FormatException("Group title is required.");
                if (string.IsNullOrWhiteSpace(item.Id)) item.Id = CustomActionStore.MakeActionId(item.Title);
                foreach (var entry in item.Actions)
                {
                    if (string.IsNullOrWhiteSpace(entry.Title) || string.IsNullOrWhiteSpace(entry.Prompt))
                    {
                        throw new FormatException("Action title and prompt are required.");
                    }
                    if (string.IsNullOrWhiteSpace(entry.Id)) entry.Id = CustomActionStore.MakeActionId(entry.Title);
                }
            }
            if (groups.Count == 0) throw new FormatException("No groups found in YAML.");
            return new CustomActionFile
            {
                SchemaVersion = CustomActionStore.CurrentSchemaVersion,
                Groups = groups.ToArray()
            };
        }

        private static CustomActionDefinition DefaultAction()
        {
            return new CustomActionDefinition
            {
                Surface = CustomActionSurface.Assistant,
                Output = "chat",
                ApplicabilityItemType = CustomActionApplicability.All,
                ApplicabilityDirection = CustomActionApplicability.All,
                AllowedTools = new string[0],
                Context = new CustomActionContext
                {
                    Source = "related_thread",
                    MessageScope = "thread",
                    FolderScope = "current_folder",
                    ReadFilter = "all",
                    TimeRange = "today",
                    IncludeFullBodies = true,
                    MaxItems = 20
                }
            };
        }

        private static void AssignGroup(CustomActionGroup group, string expression)
        {
            Split(expression, out var key, out var value);
            if (key == "id") group.Id = value;
            else if (key == "title") group.Title = value;
        }

        private static void AssignAction(CustomActionDefinition action, string expression)
        {
            Split(expression, out var key, out var value);
            if (key == "id") action.Id = value;
            else if (key == "title") action.Title = value;
            else if (key == "description") action.Description = value;
            else if (key == "prompt") action.Prompt = value;
            else if (key == "surface") action.Surface = CustomActionSurface.Normalize(value);
            else if (key == "output") action.Output = value;
            else if (key == "use_skills") action.UseSkills = string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
            else if (key == "applicability_item_type")
                action.ApplicabilityItemType = CustomActionApplicability.NormalizeItemType(value);
            else if (key == "applicability_direction")
                action.ApplicabilityDirection = CustomActionApplicability.NormalizeDirection(value);
            else if (key == "disabled") action.Disabled = string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
            else if (key == "context_json" && !string.IsNullOrWhiteSpace(value))
                action.Context = JsonConvert.DeserializeObject<CustomActionContext>(value) ?? action.Context;
            else if (key == "allowed_tools_json" && !string.IsNullOrWhiteSpace(value))
            {
                action.AllowedTools = JsonConvert.DeserializeObject<string[]>(value) ?? new string[0];
                action.AllowTools = action.AllowedTools.Length > 0;
            }
        }

        private static void Split(string expression, out string key, out string value)
        {
            var index = expression.IndexOf(':');
            if (index < 0)
            {
                key = expression.Trim();
                value = "";
                return;
            }
            key = expression.Substring(0, index).Trim();
            value = Unscalar(expression.Substring(index + 1).Trim());
        }

        private static string Scalar(string value)
        {
            return JsonConvert.SerializeObject(value ?? "");
        }

        private static string Unscalar(string value)
        {
            if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"')
            {
                return JsonConvert.DeserializeObject<string>(value);
            }
            if (value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'')
            {
                return value.Substring(1, value.Length - 2).Replace("''", "'");
            }
            return value;
        }

        private static string MatchKey(string group, string action)
        {
            return Normalize(group) + "\n" + Normalize(action);
        }

        private static string Normalize(string value)
        {
            return string.Join(" ", (value ?? "").Trim().ToLowerInvariant()
                .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
        }
    }
}
