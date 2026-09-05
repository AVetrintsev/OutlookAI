using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Skills;
using Xunit;

namespace OutlookAI.Tests.Services.Skills
{
    public sealed class SkillSessionTests : IDisposable
    {
        private readonly string _directory = Path.Combine(Path.GetTempPath(), "OutlookAI-skill-session-" + Guid.NewGuid().ToString("N"));
        private SkillStore Store => new SkillStore(Path.Combine(_directory, "skills.json"));
        private static readonly DateTime Today = new DateTime(2026, 9, 5);
        private Task<string> Call(SkillSession session, string tool, object args) => session.DispatchAsync(tool, SkillJson.Write(args), CancellationToken.None);

        [Fact]
        public async Task DisclosureIsRequiredEvenIfSkillsExist()
        {
            Store.Upsert(SkillStoreTests.Example(), 0);
            var session = new SkillSession(null, true, Store, Today);
            Assert.False(session.Available);
            Assert.Contains("skills_unavailable", await Call(session, SkillToolNames.List, new { }));
            Store.AcceptDisclosure();
            Assert.False(new SkillSession(null, false, Store, Today).Available);
        }

        [Fact]
        public async Task CatalogOnlyDisclosesNameAndDescriptionAndFiltersInactive()
        {
            var active = SkillStoreTests.Example(); active.Instructions = "PRIVATE INSTRUCTIONS";
            Store.Upsert(active, 0);
            var inactive = SkillStoreTests.Example("inactive"); inactive.Enabled = false; Store.Upsert(inactive, 0);
            var expired = SkillStoreTests.Example("expired"); expired.ValidUntil = Today.AddDays(-1); Store.Upsert(expired, 0);
            Store.AcceptDisclosure();
            var session = new SkillSession(null, true, Store, Today);
            var raw = await Call(session, SkillToolNames.List, new { query = "owner@example.com" });
            Assert.DoesNotContain("PRIVATE", raw);
            Assert.DoesNotContain("owner@example.com", raw);
            var skills = (JArray)JObject.Parse(raw)["skills"];
            Assert.Single(skills);
            Assert.Equal(new[] { "id", "name", "description" }, ((JObject)skills[0]).Properties().Select(p => p.Name));
            var loaded = await Call(session, SkillToolNames.Load, new { ids = new[] { "roles", "expired" } });
            Assert.Contains("PRIVATE INSTRUCTIONS", loaded);
            Assert.Equal(new[] { "Роли" }, session.LoadedNames);
            Assert.Equal("expired", (string)JObject.Parse(loaded)["unavailable_ids"][0]);
        }

        [Fact]
        public async Task LoadingIsBoundedAndUnknownIdsDoNotCountAsUsed()
        {
            for (var i = 0; i < 11; i++) Store.Upsert(SkillStoreTests.Example("s" + i), 0);
            Store.AcceptDisclosure();
            var session = new SkillSession(null, true, Store, Today);
            Assert.Contains("error", await Call(session, SkillToolNames.Load, new { ids = Enumerable.Range(0, 6).Select(i => "s" + i).ToArray() }));
            await Call(session, SkillToolNames.Load, new { ids = Enumerable.Range(0, 5).Select(i => "s" + i).ToArray() });
            await Call(session, SkillToolNames.Load, new { ids = Enumerable.Range(5, 5).Select(i => "s" + i).ToArray() });
            Assert.Contains("error", await Call(session, SkillToolNames.Load, new { ids = new[] { "s10" } }));
            Assert.Equal(10, session.LoadedNames.Count);
        }

        [Fact]
        public async Task AggregateSizeLimitRejectsTheWholeLoadWithoutChangingPreviouslyLoadedSkills()
        {
            for (var i = 0; i < 6; i++)
            {
                var skill = SkillStoreTests.Example("large" + i);
                skill.Name = "Правило " + i;
                skill.Instructions = new string((char)('a' + i), 30000);
                Store.Upsert(skill, 0);
            }
            Store.AcceptDisclosure();
            var session = new SkillSession(null, true, Store, Today);
            var first = JObject.Parse(await Call(session, SkillToolNames.Load, new { ids = new[] { "large0", "large1", "large2" } }));
            Assert.Null(first["error"]);
            Assert.Equal(3, ((JArray)first["skills"]).Count);

            var rejected = JObject.Parse(await Call(session, SkillToolNames.Load, new { ids = new[] { "large3", "large4", "large5" } }));
            Assert.Contains("160000", (string)rejected["error"]["message"]);
            Assert.Null(rejected["skills"]);
            Assert.Equal(new[] { "Правило 0", "Правило 1", "Правило 2" }, session.LoadedNames);

            var smaller = JObject.Parse(await Call(session, SkillToolNames.Load, new { ids = new[] { "large3" } }));
            Assert.Null(smaller["error"]);
            Assert.Equal("large3", (string)Assert.Single((JArray)smaller["skills"])["id"]);
            Assert.Equal(new[] { "Правило 0", "Правило 1", "Правило 2", "Правило 3" }, session.LoadedNames);
        }

        [Fact]
        public async Task ProposedUpdateRequiresLoadAndNeverWrites()
        {
            Store.Upsert(SkillStoreTests.Example(), 0); Store.AcceptDisclosure();
            var session = new SkillSession(null, true, Store, Today);
            var args = new { id = "roles", instructions = "Новая роль", reason = "Уточнение в письме" };
            Assert.Contains("error", await Call(session, SkillToolNames.ProposeUpdate, args));
            await Call(session, SkillToolNames.Load, new { ids = new[] { "roles" } });
            var before = File.ReadAllText(Store.Path);
            var result = JObject.Parse(await Call(session, SkillToolNames.ProposeUpdate, args));
            Assert.False((bool)result["saved"]);
            Assert.Equal("Новая роль", (string)result["skill_proposal"]["skill"]["instructions"]);
            Assert.Equal(before, File.ReadAllText(Store.Path));
        }

        [Fact]
        public async Task PinsProduceActualToolHistoryAndFreshSessionSeesChanges()
        {
            var saved = Store.Upsert(SkillStoreTests.Example(), 0); Store.AcceptDisclosure();
            var session = new SkillSession(null, true, Store, Today);
            var history = await session.LoadPinnedAsync(new[] { "roles", "roles" }, new ChatEventSink(), CancellationToken.None);
            Assert.Equal(2, history.Count);
            Assert.Equal((string)history[0]["call_id"], (string)history[1]["call_id"]);
            saved.Enabled = false; Store.Upsert(saved, saved.Revision);
            var fresh = new SkillSession(null, true, Store, Today);
            await Call(fresh, SkillToolNames.Load, new { ids = new[] { "roles" } });
            Assert.Empty(fresh.LoadedNames);
        }

        public void Dispose() { if (Directory.Exists(_directory)) Directory.Delete(_directory, true); }
    }
}
