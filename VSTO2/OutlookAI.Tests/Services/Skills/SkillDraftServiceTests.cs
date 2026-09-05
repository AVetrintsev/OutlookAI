using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Skills;
using OutlookAI.Services.Tools;
using Xunit;

namespace OutlookAI.Tests.Services.Skills
{
    public sealed class SkillDraftServiceTests
    {
        [Fact]
        public async Task DraftPreservesExistingIdentityAndOnlyUsesActualSourceMetadata()
        {
            var existing = SkillStoreTests.Example(); existing.Revision = 7; existing.Enabled = false;
            var generated = SkillStoreTests.Example("model-invented-id");
            generated.Sources = new[] { new SkillSource { MessageId = "invented" } };
            var service = new SkillDraftService((system, user, ct) => Task.FromResult(SkillJson.Write(new SkillDraft { Skill = generated })));
            var result = await service.CreateAsync("Уточни роль", new ConversationResult { Messages = new[] {
                new MessageDetail { Id = "m1", Subject = "Уточнение роли", From = "a@example.com", BodyPlaintext = "FULL PRIVATE EMAIL" }
            } }, existing, CancellationToken.None);
            Assert.Equal(existing.Id, result.Skill.Id);
            Assert.Equal(7, result.Skill.Revision);
            Assert.False(result.Skill.Enabled);
            Assert.Equal("m1", Assert.Single(result.Skill.Sources).MessageId);
            Assert.DoesNotContain("FULL PRIVATE EMAIL", SkillJson.Write(result.Skill.Sources));
        }

        [Fact]
        public async Task NoContextMeansNoEmailIsTransmitted()
        {
            var service = new SkillDraftService((system, input, ct) =>
            {
                Assert.Equal(JTokenType.Null, JObject.Parse(input)["email_context"].Type);
                Assert.Contains("Разовое поручение", system);
                return Task.FromResult("```json\n" + SkillJson.Write(new SkillDraft { Skill = SkillStoreTests.Example() }) + "\n```");
            });
            var result = await service.CreateAsync("Создай правило", null, null, CancellationToken.None);
            Assert.Equal(0, result.Skill.Revision);
            Assert.Empty(result.Skill.Sources);
        }

        [Theory]
        [InlineData("null")]
        [InlineData("not json")]
        [InlineData("{}")]
        public async Task InvalidModelDraftIsNotAccepted(string response)
        {
            var service = new SkillDraftService((system, input, ct) => Task.FromResult(response));
            await Assert.ThrowsAnyAsync<Exception>(() => service.CreateAsync("Новое правило", null, null, CancellationToken.None));
        }
    }
}
