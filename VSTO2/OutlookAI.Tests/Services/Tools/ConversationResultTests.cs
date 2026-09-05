using System;
using System.Threading;
using System.Threading.Tasks;
using Moq;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Tools
{
    public sealed class ConversationResultTests
    {
        [Fact]
        public void FallbackExplicitlyReportsMissingThreadAndAttachments()
        {
            var surface = new Mock<IOutlookSurface>();
            surface.Setup(s => s.GetCurrentSelection(true, 1)).Returns(new CurrentSelectionResult { Messages = new[] { new MessageDetail { Id = "m1" } } });
            var result = ConversationResult.ReadCurrent(surface.Object, 100, CancellationToken.None);
            Assert.True(result.Truncated);
            Assert.Single(result.Messages);
            Assert.False((bool)result.ToJson()["attachment_contents_analyzed"]);
        }

        [Fact]
        public void ProjectionOrdersByMessageDateAndPreservesTruncationFlags()
        {
            var result = new ConversationResult { Truncated = true, UnavailableCount = 2, Messages = new[] {
                new MessageDetail { Id = "new", SentAt = new DateTimeOffset(2026, 9, 5, 10, 0, 0, TimeSpan.Zero), BodyTruncated = true },
                new MessageDetail { Id = "old", ReceivedAt = new DateTimeOffset(2026, 9, 4, 10, 0, 0, TimeSpan.Zero) }
            } };
            var json = result.ToJson();
            Assert.Equal("old", (string)json["messages"][0]["id"]);
            Assert.True((bool)json["messages"][1]["body_truncated"]);
            Assert.Equal(2, (int)json["unavailable_count"]);
        }

        [Fact]
        public async Task ToolClampsLimitAndPassesExplicitMessageId()
        {
            var surface = new Mock<IOutlookSurface>();
            var conversations = surface.As<IConversationSurface>();
            conversations.Setup(s => s.ReadConversation("m1", 100, CancellationToken.None)).Returns(new ConversationResult());
            await new OutlookReadConversationTool().ExecuteAsync("{\"message_id\":\"m1\",\"max_items\":900}", surface.Object, CancellationToken.None);
            conversations.VerifyAll();
        }

        [Fact]
        public void CustomActionContextOptionsCanExcludeBodiesAndAttachmentNames()
        {
            var result = new ConversationResult { Messages = new[] { new MessageDetail {
                Id = "m1", BodyPlaintext = "PRIVATE_BODY", Attachments = new[] { new AttachmentSummary { Filename = "PRIVATE_NAME.pdf" } }
            } } };
            var json = result.ToJson(false, false).ToString();
            Assert.DoesNotContain("PRIVATE_BODY", json);
            Assert.DoesNotContain("PRIVATE_NAME", json);
            Assert.False((bool)result.ToJson(false, false)["message_bodies_included"]);
        }
    }
}
