using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services
{
    [Collection("Config")]
    public class LiteLlmChatServiceTests
    {
        public LiteLlmChatServiceTests()
        {
            Config.ResetDefaults();
            Config.LiteLlmBaseUrl = "https://litellm.local/v1";
            Config.LiteLlmApiKey = "sk-test";
            Config.Model = "company/default";
        }

        [Fact]
        public async Task ProcessEmailAsync_SendsChatCompletionRequestWithBearer()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"Hello\"}}]}\n\n"
                + "data: {\"choices\":[{\"delta\":{\"content\":\" world\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, chatHttp))
            {
                var result = await chat.ProcessEmailAsync(
                    LiteLlmChatService.ActionType.Proofread, "helo world");

                Assert.Equal("Hello world", result);
                Assert.Single(fake.Requests);
                Assert.Equal("https://litellm.local/v1/chat/completions",
                    fake.Requests[0].RequestUri.ToString());
                Assert.Equal("Bearer sk-test",
                    fake.Requests[0].Headers.Authorization.ToString());
                Assert.Contains("\"model\":\"company/default\"", fake.RequestBodies[0]);
                Assert.Contains("\"messages\":[", fake.RequestBodies[0]);
                Assert.Contains("\"stream\":true", fake.RequestBodies[0]);
            }
        }

        [Fact]
        public async Task ProcessEmailAsync_ThrowsWithBackendErrorBody_OnNonSuccess()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueText((HttpStatusCode)429, "{\"detail\":\"rate limited\"}");

            using (var credentials = new LiteLlmCredentialService())
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, chatHttp))
            {
                var ex = await Assert.ThrowsAsync<InvalidOperationException>(
                    () => chat.ProcessEmailAsync(
                        LiteLlmChatService.ActionType.Proofread, "x"));

                Assert.Contains("LiteLLM backend error", ex.Message);
                Assert.Contains("rate limited", ex.Message);
            }
        }
    }
}
