using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.CustomActions;
using OutlookAI.Services.Tools;
using OutlookAI.Tests.Helpers;
using OutlookAI.Tests.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.CustomActions
{
    [Collection("Config")]
    public sealed class CustomActionRunnerTests
    {
        public CustomActionRunnerTests()
        {
            Config.ResetDefaults();
            Config.LiteLlmBaseUrl = "https://litellm.local/v1";
            Config.LiteLlmApiKey = "sk-test";
            Config.Model = "company/default";
        }

        [Fact]
        public async Task RunAsync_CreateReply_CreatesReplyDraftAndReturnsStatus()
        {
            var surface = new RunnerSurface();
            var result = await RunAsync(surface, Action("create_reply"));

            Assert.Equal("Черновик ответа создан", result.Text);
            Assert.Equal("create_reply", result.Output);
            Assert.Equal("reply-draft", result.DraftId);
            Assert.NotNull(surface.ReplyArgs);
            Assert.Equal("m1", surface.ReplyArgs.SourceMessageId);
            Assert.Equal("LLM body", surface.ReplyArgs.BodyPlaintext);
            Assert.Null(surface.MeetingArgs);
        }

        [Fact]
        public async Task RunAsync_CreateDraftAlias_CreatesReplyDraft()
        {
            var surface = new RunnerSurface();
            var result = await RunAsync(surface, Action("create_draft"));

            Assert.Equal("Черновик ответа создан", result.Text);
            Assert.Equal("create_reply", result.Output);
            Assert.NotNull(surface.ReplyArgs);
        }

        [Fact]
        public async Task RunAsync_CreateMeeting_CreatesMeetingDraftAndReturnsStatus()
        {
            var surface = new RunnerSurface();
            var result = await RunAsync(surface, Action("create_meeting"));

            Assert.Equal("Черновик встречи создан", result.Text);
            Assert.Equal("create_meeting", result.Output);
            Assert.Equal("meeting-draft", result.DraftId);
            Assert.NotNull(surface.MeetingArgs);
            Assert.Equal("m1", surface.MeetingArgs.SourceMessageId);
            Assert.Equal("LLM body", surface.MeetingArgs.BodyPlaintext);
            Assert.Null(surface.ReplyArgs);
        }

        [Fact]
        public async Task RunAsync_CreateReplyWithoutSelectedMessage_ReturnsError()
        {
            var surface = new RunnerSurface { SelectionMessageId = null };
            var result = await RunAsync(surface, Action("create_reply"));

            Assert.Equal("create_reply", result.Output);
            Assert.Contains("Выберите или откройте письмо", result.Text);
            Assert.Null(result.DraftId);
            Assert.Null(surface.ReplyArgs);
        }

        [Fact]
        public async Task RunAsync_NotApplicableAction_ReturnsErrorBeforeCreatingDraft()
        {
            var surface = new RunnerSurface
            {
                SelectionItemType = "mail",
                SelectionDirection = "incoming"
            };
            var action = Action("create_meeting");
            action.ApplicabilityItemType = "meeting";
            action.ApplicabilityDirection = "outgoing";

            var result = await RunAsync(surface, action);

            Assert.Contains("недоступно", result.Text);
            Assert.Null(surface.MeetingArgs);
        }

        [Fact]
        public void CustomActionUiSerializer_ShowsCreateDraftAliasAsCreateReply()
        {
            var json = CustomActionUiSerializer.ToJson(Action("create_draft"));

            Assert.Equal("create_reply", (string)json["output"]);
        }

        [Fact]
        public void CustomActionUiSerializer_ProjectsApplicabilityFields()
        {
            var action = Action("chat");
            action.ApplicabilityItemType = "meeting";
            action.ApplicabilityDirection = "outgoing";

            var json = CustomActionUiSerializer.ToJson(action);

            Assert.Equal("meeting", (string)json["applicability_item_type"]);
            Assert.Equal("outgoing", (string)json["applicability_direction"]);
        }

        private static async Task<CustomActionRunResult> RunAsync(
            RunnerSurface surface,
            CustomActionDefinition action)
        {
            var fake = new FakeHttpMessageHandler();
            fake.QueueSse(HttpStatusCode.OK,
                "data: {\"choices\":[{\"delta\":{\"content\":\"LLM body\"}}]}\n\n"
                + "data: [DONE]\n\n");

            using (var credentials = new LiteLlmCredentialService())
            using (var http = new HttpClient(fake))
            using (var chat = new LiteLlmChatService(credentials, http))
            {
                var statePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".json");
                var runner = new CustomActionRunner(
                    chat,
                    surface,
                    state: new CustomActionStateStore(statePath));
                return await runner.RunAsync(action, CancellationToken.None);
            }
        }

        private static CustomActionDefinition Action(string output)
        {
            return new CustomActionDefinition
            {
                Id = "a1",
                Title = "Action",
                Prompt = "Write a reply",
                Output = output,
                Context = new CustomActionContext
                {
                    Source = "current_selection",
                    IncludeFullBodies = true,
                    MaxItems = 1
                }
            };
        }

        private sealed class RunnerSurface : MinimalSurface
        {
            public string SelectionMessageId { get; set; } = "m1";
            public string SelectionItemType { get; set; } = "mail";
            public string SelectionDirection { get; set; } = "incoming";
            public CreateReplyDraftArgs ReplyArgs { get; private set; }
            public CreateMeetingDraftArgs MeetingArgs { get; private set; }

            public override CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems)
            {
                return new CurrentSelectionResult
                {
                    Count = string.IsNullOrWhiteSpace(SelectionMessageId) ? 0 : 1,
                    Messages = string.IsNullOrWhiteSpace(SelectionMessageId)
                        ? new MessageDetail[0]
                        : new[]
                        {
                            new MessageDetail
                            {
                                Id = SelectionMessageId,
                                ItemType = SelectionItemType,
                                Direction = SelectionDirection,
                                Subject = "Subject",
                                From = "sender@example.com",
                                To = new[] { "user@example.com" },
                                Cc = new string[0],
                                BodyPlaintext = "Body"
                            }
                        }
                };
            }

            public override CreatedDraft CreateReplyDraft(CreateReplyDraftArgs args)
            {
                if (string.IsNullOrWhiteSpace(args?.SourceMessageId))
                {
                    throw new InvalidOperationException("No source message is selected.");
                }
                ReplyArgs = args;
                return new CreatedDraft
                {
                    DraftId = "reply-draft",
                    Location = "Drafts",
                    DisplayName = "Re: Subject"
                };
            }

            public override CreatedDraft CreateMeetingDraft(CreateMeetingDraftArgs args)
            {
                if (string.IsNullOrWhiteSpace(args?.SourceMessageId))
                {
                    throw new InvalidOperationException("No source message is selected.");
                }
                MeetingArgs = args;
                return new CreatedDraft
                {
                    DraftId = "meeting-draft",
                    Location = "Calendar",
                    DisplayName = "Subject"
                };
            }
        }
    }
}
