using System;
using System.IO;
using OutlookAI.Services.Skills;
using Xunit;

namespace OutlookAI.Tests.Services.Skills
{
    public sealed class SkillStoreTests : IDisposable
    {
        private readonly string _directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "OutlookAI-skills-test-" + Guid.NewGuid().ToString("N"));
        private SkillStore Store => new SkillStore(System.IO.Path.Combine(_directory, "skills.json"));
        public static SkillDefinition Example(string id = "roles") => new SkillDefinition
        {
            Id = id, Name = "Роли", Description = "Кто отвечает за проект", Instructions = "Проверяй роли по email.",
            Category = "people", People = new[] { new SkillPerson { Email = "owner@example.com", Name = "Иван", Role = "Владелец" } }
        };

        [Fact]
        public void ReadDoesNotCreateFileOrConsent()
        {
            Assert.Empty(Store.Read().Skills);
            Assert.False(Store.Read().DisclosureAccepted);
            Assert.False(Directory.Exists(_directory));
        }

        [Fact]
        public void SavesAtomicallyAndRejectsStaleRevision()
        {
            var saved = Store.Upsert(Example(), 0);
            Assert.Equal(1, saved.Revision);
            saved.Name = "Обновление";
            Assert.Throws<InvalidOperationException>(() => Store.Upsert(saved, 0));
            Assert.Equal("Роли", Store.Read().Skills[0].Name);
            Assert.Equal(2, Store.Upsert(saved, 1).Revision);
            Assert.True(File.Exists(Store.Path + ".bak"));
            Assert.Equal(1, SkillJson.Parse<SkillFile>(File.ReadAllText(Store.Path + ".bak")).Skills[0].Revision);
        }

        [Fact]
        public void ImportOnlyReplacesExplicitlyChosenConflictsAndNeverImportsConsent()
        {
            Store.Upsert(Example(), 0);
            var changed = Example(); changed.Name = "Импорт";
            var input = SkillJson.Write(new SkillFile { DisclosureAccepted = true, Skills = new[] { changed, Example("new") } });
            var incoming = Store.ParseImport(input);
            Store.ApplyImport(incoming, Store.Read().Revision, new string[0]);
            Assert.Equal(2, Store.Read().Skills.Length);
            Assert.Equal("Роли", Store.Read().Skills[0].Name);
            Assert.False(Store.Read().DisclosureAccepted);
            Store.ApplyImport(incoming, Store.Read().Revision, new[] { "roles" });
            Assert.Equal("Импорт", Store.Read().Skills[0].Name);
        }

        [Fact]
        public void ImportPreviewBecomesStaleAfterAnotherWrite()
        {
            var revision = Store.Read().Revision;
            Store.Upsert(Example(), 0);
            Assert.Throws<InvalidOperationException>(() => Store.ApplyImport(new[] { Example("new") }, revision, new string[0]));
            Assert.Single(Store.Read().Skills);
        }

        [Fact]
        public void DuplicateIdsAndEmailsAreRejected()
        {
            Assert.Throws<ArgumentException>(() => Store.ParseImport(SkillJson.Write(new SkillFile { Skills = new[] { Example(), Example() } })));
            Assert.Throws<ArgumentException>(() => Store.ParseImport(SkillJson.Write(new SkillFile { Skills = new[] { Example(), Example(" ROLES ") } })));
            var invalid = Example(); invalid.People = new[] { invalid.People[0], invalid.People[0] };
            Assert.Throws<ArgumentException>(() => Store.Upsert(invalid, 0));
            invalid = Example(); invalid.People[0].Email = "Иван";
            Assert.Throws<ArgumentException>(() => Store.Upsert(invalid, 0));
        }

        [Fact]
        public void NullArraysNormalizeOnImportAndRead()
        {
            var value = Example(); value.Triggers = null; value.People[0].Aliases = null;
            Store.Upsert(value, 0);
            Assert.Empty(Store.Read().Skills[0].Triggers);
            Assert.Empty(Store.Read().Skills[0].People[0].Aliases);
        }

        [Fact]
        public void DatesAreInclusiveAndDisabledAlwaysWins()
        {
            var value = Example(); value.ValidFrom = new DateTime(2026, 9, 5); value.ValidUntil = value.ValidFrom;
            Assert.True(value.IsActive(value.ValidFrom.Value));
            Assert.False(value.IsActive(value.ValidFrom.Value.AddDays(-1)));
            Assert.False(value.IsActive(value.ValidFrom.Value.AddDays(1)));
            value.Enabled = false;
            Assert.False(value.IsActive(value.ValidFrom.Value));
        }

        [Fact]
        public void DeleteRequiresCurrentRevision()
        {
            Store.Upsert(Example(), 0);
            Assert.Throws<InvalidOperationException>(() => Store.Delete("roles", 0));
            Store.Delete("roles", 1);
            Assert.Empty(Store.Read().Skills);
        }

        public void Dispose() { if (Directory.Exists(_directory)) Directory.Delete(_directory, true); }
    }
}
