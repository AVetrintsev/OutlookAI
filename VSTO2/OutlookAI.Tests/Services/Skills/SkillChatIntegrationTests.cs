using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Skills;
using OutlookAI.Services.TextEditing;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services.Skills
{
    [Collection("Config")]
    public sealed class SkillChatIntegrationTests : IDisposable
    {
        private readonly string _directory = Path.Combine(Path.GetTempPath(), "OutlookAI-skill-chat-" + Guid.NewGuid().ToString("N"));
        private SkillStore Store => new SkillStore(Path.Combine(_directory, "skills.json"));
        public SkillChatIntegrationTests()
        {
            Config.ResetDefaults(); Config.LiteLlmBaseUrl = "https://litellm.local/v1";
            Config.LiteLlmApiKey = "sk-test"; Config.Model = "company/default";
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task PinsAreLoadedBeforeHttpAndAttributionIsSuppressedForReplacement(bool suppress)
        {
            var skill = SkillStoreTests.Example(); skill.Instructions = "PRIVATE_ROLE_FACT";
            Store.Upsert(skill, 0); Store.AcceptDisclosure();
            var http = new FakeHttpMessageHandler();
            http.QueueSse(HttpStatusCode.OK, "data: {\"choices\":[{\"delta\":{\"content\":\"Готовый текст\"}}]}\n\ndata: [DONE]\n\n");
            using (var credentials = new LiteLlmCredentialService())
            using (var client = new HttpClient(http))
            using (var chat = new LiteLlmChatService(credentials, client, skillStore: Store))
            {
                var result = await chat.RunTurnAsync(new ConversationContext
                {
                    EnableSkills = true, PinnedSkillIds = new[] { skill.Id }, AllowedToolNames = new string[0],
                    IncludeWriteTools = false, SuppressSkillAttribution = suppress
                }, "Помоги", new FakeToolHost(), new CapturingChatEventSink(), CancellationToken.None);
                Assert.Contains("PRIVATE_ROLE_FACT", http.RequestBodies[0]);
                var request = JObject.Parse(http.RequestBodies[0]);
                Assert.Equal(SkillToolNames.All.OrderBy(name => name),
                    ((JArray)request["tools"]).Select(tool => (string)tool["function"]["name"]).OrderBy(name => name));
                if (suppress) Assert.Equal("Готовый текст", result.FinalAssistantText);
                else Assert.Contains("Использованы скиллы: Роли", result.FinalAssistantText);
            }
        }

        [Fact]
        public async Task SelectionShowsOnlyActuallyLoadedSkillNamesOutsideReplacement()
        {
            var pinned = SkillStoreTests.Example();
            var loaded = SkillStoreTests.Example("style"); loaded.Name = "Стиль";
            var catalogOnly = SkillStoreTests.Example("catalog-only"); catalogOnly.Name = "Не загружен";
            Store.Upsert(pinned, 0); Store.Upsert(loaded, 0); Store.Upsert(catalogOnly, 0);
            Store.AcceptDisclosure();
            var http = new FakeHttpMessageHandler();
            QueueTool(http, "catalog", SkillToolNames.List, new JObject());
            QueueTool(http, "load", SkillToolNames.Load, new JObject(new JProperty("ids", new JArray("style", "missing"))));
            QueueTool(http, "failed", SkillToolNames.Load, new JObject(new JProperty("ids", new JArray())));
            http.QueueSse(HttpStatusCode.OK, "data: {\"choices\":[{\"delta\":{\"content\":\"Готовый текст\"}}]}\n\ndata: [DONE]\n\n");
            var sink = new TextActionSkillEventSink();
            using (var credentials = new LiteLlmCredentialService())
            using (var client = new HttpClient(http))
            using (var chat = new LiteLlmChatService(credentials, client, skillStore: Store))
            {
                var result = await chat.RunTurnAsync(new ConversationContext
                {
                    EnableSkills = true, PinnedSkillIds = new[] { pinned.Id }, AllowedToolNames = new string[0],
                    IncludeWriteTools = false, SuppressSkillAttribution = true
                }, "Отредактируй", new FakeToolHost(), sink, CancellationToken.None);
                Assert.Equal(StopReason.Completed, result.StopReason);
                Assert.Equal("Готовый текст", result.FinalAssistantText);
                Assert.Equal(new[] { "Роли", "Стиль" }, sink.LoadedNames);
            }
        }

        private static void QueueTool(FakeHttpMessageHandler http, string id, string name, JObject args)
        {
            var call = new JObject(new JProperty("index", 0), new JProperty("id", id), new JProperty("type", "function"),
                new JProperty("function", new JObject(new JProperty("name", name),
                    new JProperty("arguments", args.ToString(Newtonsoft.Json.Formatting.None)))));
            var chunk = new JObject(new JProperty("choices", new JArray(new JObject(new JProperty("delta",
                new JObject(new JProperty("tool_calls", new JArray(call))))))));
            http.QueueSse(HttpStatusCode.OK, "data: " + chunk.ToString(Newtonsoft.Json.Formatting.None) + "\n\ndata: [DONE]\n\n");
        }

        [Fact]
        public async Task DisabledSkillsAreNotResentFromEarlierToolOutputs()
        {
            var http = new FakeHttpMessageHandler();
            http.QueueSse(HttpStatusCode.OK, "data: {\"choices\":[{\"delta\":{\"content\":\"Ответ\"}}]}\n\ndata: [DONE]\n\n");
            var context = new ConversationContext { EnableSkills = false, AllowedToolNames = new string[0] };
            context.History.Add(JObject.Parse("{\"type\":\"function_call\",\"name\":\"outlook_load_skill\",\"call_id\":\"old\",\"arguments\":\"{}\"}"));
            context.History.Add(JObject.Parse("{\"type\":\"function_call_output\",\"call_id\":\"old\",\"output\":\"EXPIRED_PRIVATE_FACT\"}"));
            using (var credentials = new LiteLlmCredentialService())
            using (var client = new HttpClient(http))
            using (var chat = new LiteLlmChatService(credentials, client, skillStore: Store))
            {
                await chat.RunTurnAsync(context, "Новый вопрос", new FakeToolHost(), null, CancellationToken.None);
                Assert.DoesNotContain("EXPIRED_PRIVATE_FACT", http.RequestBodies[0]);
                Assert.DoesNotContain("outlook_load_skill", http.RequestBodies[0]);
            }
        }
        public void Dispose() { if (Directory.Exists(_directory)) Directory.Delete(_directory, true); }
    }
}
