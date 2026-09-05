using OutlookAI.Services.Assistant;
using Xunit;

namespace OutlookAI.Tests.Services.Assistant
{
    public sealed class AssistantPlannerClientTests
    {
        [Fact]
        public void ParsePlan_CodeFencedJson_ReturnsPlan()
        {
            var plan = AssistantPlannerClient.ParsePlan(
                "```json\n{\"workflow\":\"inbox_chat\",\"intent\":\"message_sender\",\"confidence\":\"high\",\"context_mode\":\"metadata\",\"output_contract\":\"plain_answer\",\"capability_ids\":[\"tool:outlook_get_current_selection\"],\"needs_more_data\":false,\"reason\":\"metadata\"}\n```");

            Assert.Equal("inbox_chat", plan.Workflow);
            Assert.Equal("message_sender", plan.Intent);
            Assert.Equal("high", plan.Confidence);
            Assert.Equal("metadata", plan.ContextMode);
            Assert.Contains("tool:outlook_get_current_selection", plan.CapabilityIds);
        }

        [Fact]
        public void ParsePlan_NonStringFields_DoesNotThrow()
        {
            var plan = AssistantPlannerClient.ParsePlan(
                "{\"workflow\":\"inbox_chat\",\"intent\":\"create_draft\",\"confidence\":\"high\",\"context_mode\":\"snippet\",\"output_contract\":{\"id\":\"m1\"},\"capability_ids\":[\"action:reply_01\"],\"needs_more_data\":false,\"missing_data\":[],\"reason\":null}");

            Assert.Equal("create_draft", plan.Intent);
            Assert.Equal("plain_answer", plan.OutputContract);
            Assert.Equal("", plan.MissingData);
            Assert.Equal("", plan.Reason);
        }
    }
}
