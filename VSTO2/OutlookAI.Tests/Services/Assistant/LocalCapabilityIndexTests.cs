using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services;
using OutlookAI.Services.Assistant;
using OutlookAI.Services.CustomActions;
using Xunit;

namespace OutlookAI.Tests.Services.Assistant
{
    [Collection("Config")]
    public sealed class LocalCapabilityIndexTests : IDisposable
    {
        private readonly string _indexPath = Path.Combine(Path.GetTempPath(), "outlookai-capabilities-" + Guid.NewGuid().ToString("N") + ".json");
        private readonly string _actionsPath = Path.Combine(Path.GetTempPath(), "outlookai-capability-actions-" + Guid.NewGuid().ToString("N") + ".json");

        public LocalCapabilityIndexTests()
        {
            Config.ResetDefaults();
        }

        [Fact]
        public async Task Search_ReturnsSearchToolForFindMail()
        {
            var index = new LocalCapabilityIndex(
                new CapabilityCardSource(new CustomActionStore(_actionsPath)),
                new LocalHashEmbeddingProvider(),
                includeWriteTools: false,
                path: _indexPath);

            var results = await index.SearchAsync("найди письма от Андрей", new TurnSnapshot(), 5, CancellationToken.None);

            Assert.Contains(results, r => r.Card.ToolNames.Contains("outlook_search_messages"));
        }

        [Fact]
        public async Task Search_ReturnsSelectionToolForWhySelectedMessage()
        {
            var index = new LocalCapabilityIndex(
                new CapabilityCardSource(new CustomActionStore(_actionsPath)),
                new LocalHashEmbeddingProvider(),
                includeWriteTools: false,
                path: _indexPath);

            var results = await index.SearchAsync("зачем прислали это письмо", new TurnSnapshot(), 5, CancellationToken.None);

            Assert.Contains(results, r => r.Card.ToolNames.Contains("outlook_get_current_selection"));
        }

        public void Dispose()
        {
            File.Delete(_indexPath);
            File.Delete(_actionsPath);
        }

        [Fact]
        public void CardSource_IndexesLegacyAssistantActionsButNotEditorSelectionActions()
        {
            var path = Path.Combine(
                Path.GetTempPath(),
                "outlookai-capability-actions-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                var actionContext = new JObject(new JProperty("Source", "current_selection"));
                var catalog = new JObject(
                    new JProperty("SchemaVersion", CustomActionStore.CurrentSchemaVersion),
                    new JProperty("Groups", new JArray(
                        new JObject(
                            new JProperty("Id", "test_actions"),
                            new JProperty("Title", "Test actions"),
                            new JProperty("Order", 0),
                            new JProperty("Actions", new JArray(
                                new JObject(
                                    new JProperty("Id", "legacy_assistant"),
                                    new JProperty("Title", "Legacy assistant action"),
                                    new JProperty("Prompt", "Legacy prompt"),
                                    new JProperty("Context", actionContext.DeepClone()),
                                    new JProperty("Output", "chat")),
                                new JObject(
                                    new JProperty("Id", "inline_editor"),
                                    new JProperty("Title", "Inline editor action"),
                                    new JProperty("Prompt", "Inline prompt"),
                                    new JProperty("Surface", CustomActionSurface.EditorSelection),
                                    new JProperty("Context", actionContext.DeepClone()),
                                    new JProperty("Output", "replace_selection"))))))));
                File.WriteAllText(path, catalog.ToString());

                var cards = new CapabilityCardSource(new CustomActionStore(path)).LoadCards(false);

                Assert.Contains(cards, card => card.Id == "action:legacy_assistant");
                Assert.DoesNotContain(cards, card => card.Id == "action:inline_editor");
                Assert.DoesNotContain(cards, card => card.Id.StartsWith(
                    "action:text_editing_",
                    StringComparison.OrdinalIgnoreCase));
            }
            finally
            {
                try { File.Delete(path); } catch { }
            }
        }
    }
}
