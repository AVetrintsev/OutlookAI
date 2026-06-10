using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Compatibility facade for callers/tests. Tool descriptions and JSON
    /// schemas live in Prompts/tool-catalog.json; C# keeps only loading and
    /// OpenAI Chat Completions shape conversion.
    /// </summary>
    public static class ToolCatalogSchema
    {
        public static JArray BuildChatCompletionsToolsArray(bool includeWriteTools)
        {
            var responsesTools = BuildResponsesToolsArray(includeWriteTools);
            var arr = new JArray();
            foreach (var token in responsesTools)
            {
                var tool = token as JObject;
                if (tool == null)
                {
                    continue;
                }

                arr.Add(new JObject(
                    new JProperty("type", "function"),
                    new JProperty("function", new JObject(
                        new JProperty("name", (string)tool["name"] ?? ""),
                        new JProperty("description", (string)tool["description"] ?? ""),
                        new JProperty("parameters", tool["parameters"] ?? new JObject())))));
            }
            return arr;
        }

        public static JArray BuildResponsesToolsArray(bool includeWriteTools)
        {
            return ToolManifestCatalog.Default.BuildResponsesToolsArray(includeWriteTools);
        }
    }
}
