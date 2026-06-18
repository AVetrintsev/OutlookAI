using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.CustomActions;

namespace OutlookAI.Services
{
    public sealed class BuiltinActionDefinition
    {
        public string Id { get; set; }
        public string Surface { get; set; }
        public string Label { get; set; }
        public string Prompt { get; set; }
        public string Selection { get; set; }
    }

    public sealed class ActionCatalog
    {
        private readonly IReadOnlyList<BuiltinActionDefinition> _actions;
        private readonly IReadOnlyList<CustomActionGroup> _assistantGroups;
        private static readonly Lazy<ActionCatalog> DefaultInstance =
            new Lazy<ActionCatalog>(() => LoadDefault());

        private ActionCatalog(
            IReadOnlyList<BuiltinActionDefinition> actions,
            IReadOnlyList<CustomActionGroup> assistantGroups)
        {
            _actions = actions;
            _assistantGroups = assistantGroups;
        }

        public static ActionCatalog Default => DefaultInstance.Value;

        public IReadOnlyList<BuiltinActionDefinition> ForSurface(string surface)
        {
            return _actions
                .Where(a => string.Equals(a.Surface, surface, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        public IReadOnlyList<CustomActionGroup> AssistantGroups()
        {
            return _assistantGroups.Select(group => group.Clone()).ToList();
        }

        private static ActionCatalog LoadDefault()
        {
            var path = ManifestPathResolver.ResolveProjectManifest(
                Path.Combine("Prompts", "builtin-actions.json"));
            var root = JObject.Parse(File.ReadAllText(path));
            var actions = new List<BuiltinActionDefinition>();
            var groups = new List<CustomActionGroup>();

            foreach (var item in root["actions"] as JArray ?? new JArray())
            {
                var obj = item as JObject;
                var action = new BuiltinActionDefinition
                {
                    Id = (string)obj?["id"],
                    Surface = (string)obj?["surface"],
                    Label = (string)obj?["label"],
                    Prompt = (string)obj?["prompt"],
                    Selection = (string)obj?["selection"] ?? "any"
                };
                if (string.IsNullOrWhiteSpace(action.Id)
                    || string.IsNullOrWhiteSpace(action.Surface)
                    || string.IsNullOrWhiteSpace(action.Label)
                    || string.IsNullOrWhiteSpace(action.Prompt))
                {
                    throw new InvalidOperationException("Invalid action entry in " + path);
                }
                actions.Add(action);
            }

            var order = 0;
            foreach (var item in root["groups"] as JArray ?? new JArray())
            {
                var obj = item as JObject;
                var group = new CustomActionGroup
                {
                    Id = (string)obj?["id"],
                    Title = (string)obj?["title"],
                    Order = (int?)obj?["order"] ?? order++,
                    Actions = (obj?["actions"] as JArray ?? new JArray())
                        .OfType<JObject>()
                        .Select(ParseCustomAction)
                        .ToArray()
                };
                if (string.IsNullOrWhiteSpace(group.Id)
                    || string.IsNullOrWhiteSpace(group.Title)
                    || group.Actions.Length == 0)
                {
                    throw new InvalidOperationException("Invalid action group in " + path);
                }
                groups.Add(group);
            }

            return new ActionCatalog(actions, groups);
        }

        private static CustomActionDefinition ParseCustomAction(JObject obj)
        {
            var action = new CustomActionDefinition
            {
                Id = (string)(obj?["id"]),
                Title = (string)(obj?["title"]),
                Description = (string)(obj?["description"]) ?? "",
                Prompt = (string)(obj?["prompt"]),
                Output = (string)(obj?["output"]) ?? "chat",
                AllowTools = (bool?)(obj?["allow_tools"]) ?? false,
                AllowedTools = (obj?["allowed_tools"] as JArray)?.Values<string>().ToArray()
                    ?? new string[0],
                ApplicabilityItemType = CustomActionApplicability.NormalizeItemType(
                    (string)(obj?["applicability_item_type"])),
                ApplicabilityDirection = CustomActionApplicability.NormalizeDirection(
                    (string)(obj?["applicability_direction"])),
                Context = obj?["context"]?.ToObject<CustomActionContext>()
                    ?? new CustomActionContext
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
            if (string.IsNullOrWhiteSpace(action.Id)
                || string.IsNullOrWhiteSpace(action.Title)
                || string.IsNullOrWhiteSpace(action.Prompt))
            {
                throw new InvalidOperationException("Invalid action in grouped catalog.");
            }
            return action;
        }
    }
}
