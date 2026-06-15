using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Services
{
    [Collection("Config")]
    public class LiteLlmChatServiceMultiRoundTests
    {
        public LiteLlmChatServiceMultiRoundTests()
        {
            Config.ResetDefaults();
            Config.LiteLlmBaseUrl = "https://litellm.local/v1";
            Config.LiteLlmApiKey = "sk-test";
            Config.Model = "company/default";
        }

        [Fact]
        public async Task RunTurnAsync_SingleRound_NoToolCalls_ReturnsCompleted()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"hello\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, chatHttp))
            {
                var ctx = new ConversationContext { SystemInstructions = "Be brief." };
                var sink = new CapturingChatEventSink();
                var tools = new FakeToolHost();

                var result = await chat.RunTurnAsync(ctx, "say hi", tools, sink, CancellationToken.None);

                Assert.Equal(StopReason.Completed, result.StopReason);
                Assert.Equal("hello", result.FinalAssistantText);
                Assert.Equal(1, result.RoundsUsed);
                Assert.Equal("hello", sink.StreamedText.ToString());
                Assert.Single(result.AppendedItems);
                Assert.Equal(2, ctx.History.Count);
                Assert.Contains("\"tools\":[", fake.RequestBodies[0]);
                Assert.Contains("\"function\":{\"name\":\"outlook_get_current_compose_state\"", fake.RequestBodies[0]);
                Assert.DoesNotContain("parallel_tool_calls", fake.RequestBodies[0]);
            }
        }

        [Fact]
        public async Task RunTurnAsync_ToolCall_DispatchesAndAppendsToolOutput_ThenCompletes()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"call_1\",\"type\":\"function\",\"function\":{\"name\":\"outlook_get_current_compose_state\",\"arguments\":\"{}\"}}]}}]}\n\n"
                + "data: [DONE]\n\n");
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"Subject was X\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, chatHttp))
            {
                var ctx = new ConversationContext { SystemInstructions = "Be brief." };
                var sink = new CapturingChatEventSink();
                var tools = new FakeToolHost();
                tools.Queue("outlook_get_current_compose_state", "{\"subject\":\"X\"}");

                var result = await chat.RunTurnAsync(ctx, "what's the subject", tools, sink, CancellationToken.None);

                Assert.Equal(StopReason.Completed, result.StopReason);
                Assert.Equal(2, result.RoundsUsed);
                Assert.Equal("Subject was X", result.FinalAssistantText);
                Assert.Single(tools.Calls);
                Assert.Equal("outlook_get_current_compose_state", tools.Calls[0].Name);
                Assert.Equal("{}", tools.Calls[0].ArgsJson);
                Assert.Equal("call_1", sink.ToolStarts[0].CallId);
                Assert.Equal("function_call", (string)ctx.History[1]["type"]);
                Assert.Equal("function_call_output", (string)ctx.History[2]["type"]);
                Assert.Contains("\"role\":\"tool\"", fake.RequestBodies[1]);
                Assert.Contains("\"tool_call_id\":\"call_1\"", fake.RequestBodies[1]);
            }
        }

        [Fact]
        public async Task RunTurnAsync_TextFunctionJson_DispatchesWithoutShowingJson()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"{ \\\"function\\\": \\\"outlook_get_current_selection\\\", \\\"parameters\\\": { \\\"include_full_bodies\\\": true, \\\"max_items\\\": 10 } }\"}}]}\n\n"
                + "data: [DONE]\n\n");
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"Выбранное сообщение: X\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, chatHttp))
            {
                var ctx = new ConversationContext { SystemInstructions = "Be brief." };
                var sink = new CapturingChatEventSink();
                var tools = new FakeToolHost();
                tools.Queue("outlook_get_current_selection", "{\"items\":[{\"subject\":\"X\"}]}");

                var result = await chat.RunTurnAsync(ctx, "summarize selected", tools, sink, CancellationToken.None);

                Assert.Equal(StopReason.Completed, result.StopReason);
                Assert.Equal(2, result.RoundsUsed);
                Assert.Equal("Выбранное сообщение: X", result.FinalAssistantText);
                Assert.Equal("Выбранное сообщение: X", sink.StreamedText.ToString());
                Assert.Single(tools.Calls);
                Assert.Equal("outlook_get_current_selection", tools.Calls[0].Name);
                Assert.Equal("{\"include_full_bodies\":true,\"max_items\":10}", tools.Calls[0].ArgsJson);
                Assert.DoesNotContain("\"function\"", result.FinalAssistantText);
            }
        }

        [Fact]
        public async Task RunTurnAsync_RawToolResultJson_ReturnsLocalSummaryInsteadOfJson()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"PDF\\n{ \\\"folder\\\": \\\"Входящие\\\", \\\"count\\\": 1, \\\"messages\\\": [ { \\\"subject\\\": \\\"Хранилище iCloud заполнено\\\", \\\"from\\\": \\\"iCloud <noreply@email.apple.com>\\\", \\\"received_at\\\": \\\"2026-06-08T15:39:08.000+03:00\\\", \\\"body_plaintext\\\": \\\"Хранилище iCloud заполнено. Необходимо перейти на тарифный план iCloud+ или освободить место.\\\" } ] }\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, chatHttp))
            {
                var result = await chat.RunTurnAsync(
                    new ConversationContext(),
                    "summarize selected",
                    new FakeToolHost(),
                    new CapturingChatEventSink(),
                    CancellationToken.None);

                Assert.Equal(StopReason.Completed, result.StopReason);
                Assert.Contains("Сводка выбранной переписки", result.FinalAssistantText);
                Assert.Contains("Хранилище iCloud заполнено", result.FinalAssistantText);
                Assert.DoesNotContain("\"messages\"", result.FinalAssistantText);
                Assert.DoesNotContain("body_plaintext", result.FinalAssistantText);
            }
        }

        [Fact]
        public async Task RunTurnAsync_StructuredAnswerJson_UnwrapsToPlainText()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"PDF\\n{ \\\"The thread is about an order.\\\": \\\"The thread is about an order.\\\" }\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, chatHttp))
            {
                var sink = new CapturingChatEventSink();

                var result = await chat.RunTurnAsync(
                    new ConversationContext(),
                    "what is this thread about?",
                    new FakeToolHost(),
                    sink,
                    CancellationToken.None);

                Assert.Equal(StopReason.Completed, result.StopReason);
                Assert.Equal("The thread is about an order.", result.FinalAssistantText);
                Assert.Equal("The thread is about an order.", sink.StreamedText.ToString());
                Assert.DoesNotContain("PDF", result.FinalAssistantText);
                Assert.DoesNotContain("{", result.FinalAssistantText);
            }
        }

        [Fact]
        public async Task RunTurnAsync_StreamedToolArguments_AreAccumulated()
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"id\":\"call_2\",\"type\":\"function\",\"function\":{\"name\":\"outlook_count_messages\",\"arguments\":\"{\\\"query\\\"\"}}]}}]}\n\n"
                + "data: {\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"function\":{\"arguments\":\":\\\"EIN\\\"}\"}}]}}]}\n\n"
                + "data: [DONE]\n\n");
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"3\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var chatHttp = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, chatHttp))
            {
                var tools = new FakeToolHost();
                tools.Queue("outlook_count_messages", "{\"count\":3}");

                await chat.RunTurnAsync(new ConversationContext(), "count EIN", tools, new CapturingChatEventSink(), CancellationToken.None);

                Assert.Single(tools.Calls);
                Assert.Equal("{\"query\":\"EIN\"}", tools.Calls[0].ArgsJson);
            }
        }

        [Fact]
        public void FormatTraceArgs_TruncatesAtFiveHundredCharacters()
        {
            var input = new string('x', 501);

            var formatted = LiteLlmChatService.FormatTraceArgs(input);

            Assert.Equal(503, formatted.Length);
            Assert.EndsWith("...", formatted);
            Assert.Equal(new string('x', 500), formatted.Substring(0, 500));
        }
    }
}
