using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;
using OutlookAI.Services;

namespace OutlookAI.Services.Assistant
{
    public sealed class WorkflowManifestCatalog
    {
        private static readonly Lazy<WorkflowManifestCatalog> DefaultInstance =
            new Lazy<WorkflowManifestCatalog>(() => LoadDefault());

        private readonly Dictionary<string, WorkflowManifest> _workflows;

        private WorkflowManifestCatalog(IEnumerable<WorkflowManifest> workflows)
        {
            _workflows = (workflows ?? new WorkflowManifest[0])
                .Where(workflow => !string.IsNullOrWhiteSpace(workflow.Id))
                .ToDictionary(workflow => workflow.Id, StringComparer.OrdinalIgnoreCase);
        }

        public static WorkflowManifestCatalog Default => DefaultInstance.Value;

        public WorkflowManifest Get(string id)
        {
            if (!string.IsNullOrWhiteSpace(id) && _workflows.TryGetValue(id, out var workflow))
            {
                return workflow;
            }
            return _workflows.TryGetValue("inbox_chat", out var fallback)
                ? fallback
                : new WorkflowManifest { Id = "inbox_chat" };
        }

        private static WorkflowManifestCatalog LoadDefault()
        {
            var path = ManifestPathResolver.ResolveProjectManifest(
                Path.Combine("Prompts", "workflow-manifests.json"));
            var root = JObject.Parse(File.ReadAllText(path));
            var workflows = (root["workflows"] as JArray)?.OfType<JObject>()
                .Select(obj => new WorkflowManifest
                {
                    Id = (string)obj["id"] ?? "",
                    Surface = (string)obj["surface"] ?? "",
                    DefaultOutputContract = (string)obj["default_output_contract"] ?? "plain_answer",
                    DefaultContextMode = (string)obj["default_context_mode"] ?? "metadata",
                    AllowedCapabilityGroups = ((obj["allowed_capability_groups"] as JArray)?.Values<string>()
                        ?? Enumerable.Empty<string>()).ToArray()
                })
                .ToArray()
                ?? new WorkflowManifest[0];
            return new WorkflowManifestCatalog(workflows);
        }
    }

    public sealed class WorkflowManifest
    {
        public string Id { get; set; }
        public string Surface { get; set; }
        public string DefaultOutputContract { get; set; }
        public string DefaultContextMode { get; set; }
        public string[] AllowedCapabilityGroups { get; set; } = new string[0];
    }
}
