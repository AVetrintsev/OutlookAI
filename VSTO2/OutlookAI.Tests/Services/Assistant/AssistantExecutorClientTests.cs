using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.Assistant;
using OutlookAI.Services.Chat;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.Assistant
{
    [Collection("Config")]
    public sealed class AssistantExecutorClientTests
    {
        public AssistantExecutorClientTests()
        {
            Config.ResetDefaults();
            Config.LiteLlmBaseUrl = "https://litellm.local/v1";
            Config.LiteLlmApiKey = "sk-test";
            Config.Model = "company/default";
        }

        [Fact]
        public async Task ExecuteAsync_StripsAssistantRolePrefixFromTurnResult()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"### Assistant:\\nОтвет.\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var http = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, http))
            {
                var executor = new AssistantExecutorClient(chat, new FakeToolHost());

                var result = await executor.ExecuteAsync(
                    new TurnRequest { UserText = "test" },
                    new ContextBundle(),
                    new ValidatedTurnPlan
                    {
                        Plan = new TurnPlan(),
                        ContextMode = "metadata",
                        AllowedToolNames = new string[0]
                    },
                    new CapturingChatEventSink(),
                    CancellationToken.None);

                Assert.Equal("Ответ.", result.TurnResult.FinalAssistantText);
                Assert.Equal("Ответ.", result.Envelope.AnswerMarkdown);
            }
        }
    }
}
