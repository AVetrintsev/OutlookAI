using System;
using OutlookAI.Services.Assistant;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Assistant
{
    public sealed class AssistantContextResolverTests
    {
        [Fact]
        public void Resolve_MetadataMode_OmitsBody()
        {
            var bundle = new AssistantContextResolver().Resolve(
                new TurnRequest(),
                Snapshot("Latest line.\r\n\r\nFrom: Old\r\nQuoted"),
                new ValidatedTurnPlan { ContextMode = "metadata", Plan = new TurnPlan() });

            Assert.Null(bundle.Json["selected_item"]["body_snippet"]);
        }

        [Fact]
        public void Resolve_SnippetMode_StripsQuotedHistory()
        {
            var bundle = new AssistantContextResolver().Resolve(
                new TurnRequest(),
                Snapshot("Latest line.\r\n\r\nFrom: Old\r\nQuoted"),
                new ValidatedTurnPlan { ContextMode = "snippet", Plan = new TurnPlan() });

            Assert.Equal("Latest line.", (string)bundle.Json["selected_item"]["body_snippet"]);
        }

        private static TurnSnapshot Snapshot(string body)
        {
            return new TurnSnapshot
            {
                Selection = new CurrentSelectionResult
                {
                    Count = 1,
                    Messages = new[]
                    {
                        new MessageDetail
                        {
                            Id = "m1",
                            Subject = "S",
                            From = "A <a@example.com>",
                            BodyPlaintext = body,
                            ReceivedAt = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)
                        }
                    }
                }
            };
        }
    }
}
