using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.CustomActions;
using OutlookAI.Services.Tools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.CustomActions
{
    [Collection("Config")]
    public sealed class CustomActionRecommendationServiceTests
    {
        [Fact]
        public async Task RecommendAsync_FiltersUnknownAndDuplicateIds()
        {
            Config.ResetDefaults();
            Config.LiteLlmBaseUrl = "https://litellm.local/v1";
            Config.LiteLlmApiKey = "sk-test";
            Config.Model = "company/default";
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"{\\\"action_ids\\\":[\\\"a2\\\",\\\"missing\\\",\\\"a2\\\",\\\"a1\\\"]}\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var http = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, http))
            {
                var service = new CustomActionRecommendationService(chat);
                var result = await service.RecommendAsync(
                    new CurrentSelectionResult
                    {
                        Messages = new[]
                        {
                            new MessageDetail { Id = "m1", Subject = "Тема", From = "a@b", BodyPlaintext = "Текст" }
                        }
                    },
                    new[]
                    {
                        new CustomActionGroup
                        {
                            Id = "g",
                            Title = "Группа",
                            Actions = new[] { Action("a1"), Action("a2"), Action("a2") }
                        }
                    },
                    CancellationToken.None);

                Assert.Equal(new[] { "a2", "a1" }, result);
                Assert.Contains("a1", fake.RequestBodies[0]);
                Assert.Contains("Тема", fake.RequestBodies[0]);
            }
        }

        [Fact]
        public void CacheKey_UsesSelectedMessageIds()
        {
            var selection = new CurrentSelectionResult
            {
                Messages = new List<MessageDetail>
                {
                    new MessageDetail { Id = "one" },
                    new MessageDetail { Id = "two" }
                }
            };

            Assert.Equal("one|two", CustomActionRecommendationService.CacheKey(selection));
        }

        private static CustomActionDefinition Action(string id)
        {
            return new CustomActionDefinition
            {
                Id = id,
                Title = id,
                Prompt = id,
                Context = new CustomActionContext { Source = "current_selection" }
            };
        }
    }
}
