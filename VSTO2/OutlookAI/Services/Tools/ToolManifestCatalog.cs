using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    public sealed class ToolManifestCatalog
    {
        private readonly JArray _tools;
        private static readonly Lazy<ToolManifestCatalog> DefaultInstance =
            new Lazy<ToolManifestCatalog>(() => LoadDefault());

        private ToolManifestCatalog(JArray tools)
        {
            _tools = tools;
        }

        public static ToolManifestCatalog Default => DefaultInstance.Value;

        public JArray BuildResponsesToolsArray(bool includeWriteTools)
        {
            var arr = new JArray();
            foreach (var item in _tools.OfType<JObject>())
            {
                var name = (string)item["name"] ?? "";
                var isWrite = IsWriteTool(name);
                if (isWrite && !includeWriteTools)
                {
                    continue;
                }
                if (isWrite && !IsWriteToolEnabled(name))
                {
                    continue;
                }
                arr.Add((JObject)item.DeepClone());
            }
            return arr;
        }

        private static ToolManifestCatalog LoadDefault()
        {
            var path = ManifestPathResolver.ResolveProjectManifest(
                Path.Combine("Prompts", "tool-catalog.json"));
            var root = JObject.Parse(File.ReadAllText(path));
            var tools = root["tools"] as JArray;
            if (tools == null || tools.Count == 0)
            {
                throw new InvalidOperationException("No tools defined in " + path);
            }

            foreach (var tool in tools.OfType<JObject>())
            {
                if (string.IsNullOrWhiteSpace((string)tool["name"])
                    || string.IsNullOrWhiteSpace((string)tool["description"])
                    || tool["parameters"] == null)
                {
                    throw new InvalidOperationException("Invalid tool entry in " + path);
                }
            }

            return new ToolManifestCatalog(tools);
        }

        private static bool IsWriteTool(string name)
        {
            return name == "outlook_create_draft"
                || name == "outlook_mark_as_read"
                || name == "outlook_flag_message"
                || name == "outlook_set_category";
        }

        private static bool IsWriteToolEnabled(string name)
        {
            var enabled = Config.EnabledWriteTools
                ?? new System.Collections.Generic.HashSet<string>(Config.AllWriteTools);
            return enabled.Contains(name);
        }
    }
}
