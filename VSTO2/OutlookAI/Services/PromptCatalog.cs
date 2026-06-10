using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services
{
    public sealed class PromptCatalog
    {
        private readonly Dictionary<string, string> _prompts;
        private static readonly Lazy<PromptCatalog> DefaultInstance =
            new Lazy<PromptCatalog>(() => LoadDefault());

        private PromptCatalog(Dictionary<string, string> prompts)
        {
            _prompts = prompts;
        }

        public static PromptCatalog Default => DefaultInstance.Value;

        public string Get(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("Prompt id is required.", nameof(id));
            }

            if (_prompts.TryGetValue(id, out var text) && !string.IsNullOrWhiteSpace(text))
            {
                return text;
            }

            throw new InvalidOperationException("Prompt '" + id + "' is not defined.");
        }

        private static PromptCatalog LoadDefault()
        {
            var path = ManifestPathResolver.ResolveProjectManifest(
                Path.Combine("Prompts", "builtin-prompts.json"));
            var root = JObject.Parse(File.ReadAllText(path));
            var prompts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in root["prompts"] as JArray ?? new JArray())
            {
                var obj = item as JObject;
                var id = (string)obj?["id"];
                var text = (string)obj?["text"];
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(text))
                {
                    throw new InvalidOperationException("Invalid prompt entry in " + path);
                }
                prompts[id.Trim()] = text;
            }

            return new PromptCatalog(prompts);
        }
    }
}
