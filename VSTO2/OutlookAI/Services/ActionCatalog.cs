using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;

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
        private static readonly Lazy<ActionCatalog> DefaultInstance =
            new Lazy<ActionCatalog>(() => LoadDefault());

        private ActionCatalog(IReadOnlyList<BuiltinActionDefinition> actions)
        {
            _actions = actions;
        }

        public static ActionCatalog Default => DefaultInstance.Value;

        public IReadOnlyList<BuiltinActionDefinition> ForSurface(string surface)
        {
            return _actions
                .Where(a => string.Equals(a.Surface, surface, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        private static ActionCatalog LoadDefault()
        {
            var path = ManifestPathResolver.ResolveProjectManifest(
                Path.Combine("Prompts", "builtin-actions.json"));
            var root = JObject.Parse(File.ReadAllText(path));
            var actions = new List<BuiltinActionDefinition>();

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

            return new ActionCatalog(actions);
        }
    }
}
