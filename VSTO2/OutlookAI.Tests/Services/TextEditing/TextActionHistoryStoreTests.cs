using System;
using System.IO;
using System.Linq;
using OutlookAI.Services.TextEditing;
using Xunit;

namespace OutlookAI.Tests.Services.TextEditing
{
    public sealed class TextActionHistoryStoreTests : IDisposable
    {
        private readonly string _directory;
        private readonly string _path;

        public TextActionHistoryStoreTests()
        {
            _directory = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "outlookai-text-history-tests",
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
            _path = System.IO.Path.Combine(_directory, "history.json");
        }

        [Fact]
        public void AppendApplied_PersistsMetadataPreviewAndHashesWithoutFullText()
        {
            var original = "Visible original " + new string('x', 300) + " ORIGINAL_SECRET_TAIL";
            var replacement = "Visible replacement " + new string('y', 300) + " REPLACEMENT_SECRET_TAIL";
            var at = DateTimeOffset.Parse("2026-07-17T10:15:00+03:00");
            var store = new TextActionHistoryStore(_path);

            var appended = store.AppendApplied(
                "formal",
                "Сформулировать официально",
                "mail",
                original,
                replacement,
                at);
            var loaded = Assert.Single(store.ReadAll());
            var json = File.ReadAllText(_path);

            Assert.Equal(appended.Id, loaded.Id);
            Assert.Equal(TextActionHistoryStore.AppliedStatus, loaded.Status);
            Assert.Equal(at.ToUniversalTime(), loaded.CreatedAtUtc);
            Assert.Equal(TextActionHistoryStore.ComputeSha256(original), loaded.OriginalSha256);
            Assert.Equal(TextActionHistoryStore.ComputeSha256(replacement), loaded.ReplacementSha256);
            Assert.True(loaded.OriginalPreview.Length <= TextActionHistoryStore.PreviewMaxCharacters);
            Assert.True(loaded.ReplacementPreview.Length <= TextActionHistoryStore.PreviewMaxCharacters);
            Assert.DoesNotContain("ORIGINAL_SECRET_TAIL", json);
            Assert.DoesNotContain("REPLACEMENT_SECRET_TAIL", json);
            Assert.DoesNotContain(original, json);
            Assert.DoesNotContain(replacement, json);
        }

        [Fact]
        public void ReadAll_MalformedJson_ReturnsEmptyAndNextAppendRepairsFile()
        {
            File.WriteAllText(_path, "{ definitely-not-json");
            var store = new TextActionHistoryStore(_path);

            Assert.Empty(store.ReadAll());

            store.AppendApplied("shorten", "Сократить", "task", "before", "after");

            var entry = Assert.Single(store.ReadAll());
            Assert.Equal("shorten", entry.ActionId);
            Assert.Contains("\"schema_version\": 1", File.ReadAllText(_path));
        }

        [Fact]
        public void MarkUndone_ChangesStatusOnceAndPersistsTimestamp()
        {
            var store = new TextActionHistoryStore(_path);
            var entry = store.AppendApplied("clarify", "Уточнить", "meeting", "a", "b");
            var undoneAt = DateTimeOffset.Parse("2026-07-17T12:30:00Z");

            Assert.True(store.MarkUndone(entry.Id, undoneAt));
            Assert.False(store.MarkUndone(entry.Id, undoneAt.AddMinutes(1)));

            var loaded = Assert.Single(new TextActionHistoryStore(_path).ReadAll());
            Assert.Equal(TextActionHistoryStore.UndoneStatus, loaded.Status);
            Assert.Equal(undoneAt, loaded.UndoneAtUtc);
        }

        [Fact]
        public void AppendApplied_KeepsOnlyNewestTwoHundredEntries()
        {
            var store = new TextActionHistoryStore(_path);
            for (var i = 0; i < TextActionHistoryStore.MaxEntries + 5; i++)
            {
                store.AppendApplied("action-" + i, "Action " + i, "mail", "old", "new");
            }

            var entries = store.ReadAll();

            Assert.Equal(TextActionHistoryStore.MaxEntries, entries.Count);
            Assert.Equal("action-5", entries[0].ActionId);
            Assert.Equal("action-204", entries[entries.Count - 1].ActionId);
        }

        [Fact]
        public void ReadAll_ReturnsDefensiveCopies()
        {
            var store = new TextActionHistoryStore(_path);
            store.AppendApplied("formal", "Formal", "mail", "old", "new");

            var firstRead = Assert.Single(store.ReadAll());
            firstRead.Status = "tampered";

            Assert.Equal(
                TextActionHistoryStore.AppliedStatus,
                Assert.Single(store.ReadAll()).Status);
        }

        [Fact]
        public void AtomicWrites_LeaveNoTemporaryFiles()
        {
            var store = new TextActionHistoryStore(_path);
            var entry = store.AppendApplied("formal", "Formal", "mail", "old", "new");
            store.MarkUndone(entry.Id);

            Assert.Empty(Directory.GetFiles(_directory, "*.tmp"));
        }

        [Fact]
        public void ComputeSha256_IsStableLowercaseHex()
        {
            Assert.Equal(
                "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
                TextActionHistoryStore.ComputeSha256("abc"));
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_directory)) Directory.Delete(_directory, true);
            }
            catch
            {
                // Best-effort test cleanup.
            }
        }
    }
}
