using System.Linq;
using OutlookAI.Services.Assistant;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Assistant
{
    public sealed class AssistantPlanValidatorTests
    {
        [Fact]
        public void Validate_MetadataIntent_UsesNoTools()
        {
            var plan = new TurnPlan
            {
                Intent = "message_sender",
                ContextMode = "metadata",
                CapabilityIds = new[] { "tool:outlook_get_current_selection" }
            };
            var snapshot = SnapshotWithCapabilities();

            var result = new AssistantPlanValidator().Validate(plan, new TurnRequest(), snapshot, includeWriteTools: true);

            Assert.Equal("metadata", result.ContextMode);
            Assert.Empty(result.AllowedToolNames);
        }

        [Fact]
        public void Validate_ContentIntent_UsesSnippetAndSelectionTool()
        {
            var plan = new TurnPlan
            {
                Intent = "content_question",
                ContextMode = "metadata",
                CapabilityIds = new[] { "tool:outlook_get_current_selection" }
            };
            var snapshot = SnapshotWithCapabilities();

            var result = new AssistantPlanValidator().Validate(plan, new TurnRequest(), snapshot, includeWriteTools: true);

            Assert.Equal("snippet", result.ContextMode);
            Assert.Contains("outlook_get_current_selection", result.AllowedToolNames);
        }

        [Fact]
        public void Validate_ContentIntentWithoutSelection_AsksForData()
        {
            var plan = new TurnPlan { Intent = "content_question", ContextMode = "snippet" };

            var result = new AssistantPlanValidator().Validate(plan, new TurnRequest(), new TurnSnapshot(), includeWriteTools: true);

            Assert.True(result.NeedsMoreData);
            Assert.Empty(result.AllowedToolNames);
        }

        [Fact]
        public void Validate_CreateDraftIntent_AddsDraftTool()
        {
            var plan = new TurnPlan
            {
                Intent = "create_draft",
                ContextMode = "snippet",
                CapabilityIds = new[] { "action:reply_01" }
            };
            var snapshot = SnapshotWithCapabilities(new CapabilitySearchResult
            {
                Card = new CapabilityCard
                {
                    Id = "action:reply_01",
                    Type = "action",
                    Group = "custom.action",
                    Risk = "read_only",
                    ToolNames = new string[0]
                },
                Score = 1
            });

            var result = new AssistantPlanValidator().Validate(plan, new TurnRequest(), snapshot, includeWriteTools: true);

            Assert.Equal("snippet", result.ContextMode);
            Assert.Contains("outlook_create_draft", result.AllowedToolNames);
            Assert.True(result.IncludeWriteTools);
        }

        private static TurnSnapshot SnapshotWithCapabilities(params CapabilitySearchResult[] extraCapabilities)
        {
            var capabilities = new[]
            {
                new CapabilitySearchResult
                {
                    Card = new CapabilityCard
                    {
                        Id = "tool:outlook_get_current_selection",
                        Type = "tool",
                        Group = "selection.read",
                        Risk = "read_only",
                        ToolNames = new[] { "outlook_get_current_selection" }
                    },
                    Score = 1
                }
            }.Concat(extraCapabilities ?? new CapabilitySearchResult[0]).ToArray();
            return new TurnSnapshot
            {
                Selection = new CurrentSelectionResult
                {
                    Count = 1,
                    Messages = new[]
                    {
                        new MessageDetail { Id = "m1", ItemType = "mail", Subject = "S" }
                    }
                },
                Capabilities = capabilities
            };
        }
    }
}
