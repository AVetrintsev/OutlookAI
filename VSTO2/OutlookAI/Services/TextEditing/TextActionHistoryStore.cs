using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace OutlookAI.Services.TextEditing
{
    public sealed class TextActionHistoryEntry
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("created_at_utc")]
        public DateTimeOffset CreatedAtUtc { get; set; }

        [JsonProperty("undone_at_utc", NullValueHandling = NullValueHandling.Ignore)]
        public DateTimeOffset? UndoneAtUtc { get; set; }

        [JsonProperty("action_id")]
        public string ActionId { get; set; }

        [JsonProperty("action_title")]
        public string ActionTitle { get; set; }

        [JsonProperty("item_type")]
        public string ItemType { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }

        [JsonProperty("original_preview")]
        public string OriginalPreview { get; set; }

        [JsonProperty("replacement_preview")]
        public string ReplacementPreview { get; set; }

        [JsonProperty("original_sha256")]
        public string OriginalSha256 { get; set; }

        [JsonProperty("replacement_sha256")]
        public string ReplacementSha256 { get; set; }

        internal TextActionHistoryEntry Clone()
        {
            return new TextActionHistoryEntry
            {
                Id = Id,
                CreatedAtUtc = CreatedAtUtc,
                UndoneAtUtc = UndoneAtUtc,
                ActionId = ActionId,
                ActionTitle = ActionTitle,
                ItemType = ItemType,
                Status = Status,
                OriginalPreview = OriginalPreview,
                ReplacementPreview = ReplacementPreview,
                OriginalSha256 = OriginalSha256,
                ReplacementSha256 = ReplacementSha256
            };
        }
    }

    internal sealed class TextActionHistoryDocument
    {
        [JsonProperty("schema_version")]
        public int SchemaVersion { get; set; }

        [JsonProperty("entries")]
        public List<TextActionHistoryEntry> Entries { get; set; }
    }

    /// <summary>
    /// Bounded metadata-only audit history for selection replacements. Full
    /// original/replacement bodies are intentionally never serialized.
    /// </summary>
    public sealed class TextActionHistoryStore
    {
        public const int CurrentSchemaVersion = 1;
        public const int MaxEntries = 200;
        public const int PreviewMaxCharacters = 160;
        public const string AppliedStatus = "applied";
        public const string UndoneStatus = "undone";

        private static readonly object GlobalGate = new object();

        public string Path { get; }

        public TextActionHistoryStore()
            : this(OutlookAI.Services.ManifestPathResolver.AppDataFile("text-action-history.json"))
        {
        }

        public TextActionHistoryStore(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("History path is required.", nameof(path));
            }
            Path = path;
        }

        public IReadOnlyList<TextActionHistoryEntry> ReadAll()
        {
            lock (GlobalGate)
            {
                return LoadUnsafe().Entries
                    .Select(entry => entry.Clone())
                    .ToArray();
            }
        }

        public TextActionHistoryEntry AppendApplied(
            string actionId,
            string actionTitle,
            string itemType,
            string originalText,
            string replacementText,
            DateTimeOffset? timestamp = null)
        {
            var entry = new TextActionHistoryEntry
            {
                Id = Guid.NewGuid().ToString("N"),
                CreatedAtUtc = (timestamp ?? DateTimeOffset.UtcNow).ToUniversalTime(),
                ActionId = actionId ?? "",
                ActionTitle = actionTitle ?? "",
                ItemType = itemType ?? "",
                Status = AppliedStatus,
                OriginalPreview = BuildPreview(originalText),
                ReplacementPreview = BuildPreview(replacementText),
                OriginalSha256 = ComputeSha256(originalText),
                ReplacementSha256 = ComputeSha256(replacementText)
            };

            lock (GlobalGate)
            {
                var document = LoadUnsafe();
                document.Entries.Add(entry);
                while (document.Entries.Count > MaxEntries)
                {
                    document.Entries.RemoveAt(0);
                }
                SaveUnsafe(document);
            }

            return entry.Clone();
        }

        public bool MarkUndone(string entryId, DateTimeOffset? timestamp = null)
        {
            if (string.IsNullOrWhiteSpace(entryId)) return false;
            lock (GlobalGate)
            {
                var document = LoadUnsafe();
                var entry = document.Entries.FirstOrDefault(candidate =>
                    string.Equals(candidate.Id, entryId, StringComparison.OrdinalIgnoreCase));
                if (entry == null
                    || !string.Equals(entry.Status, AppliedStatus, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                entry.Status = UndoneStatus;
                entry.UndoneAtUtc = (timestamp ?? DateTimeOffset.UtcNow).ToUniversalTime();
                SaveUnsafe(document);
                return true;
            }
        }

        public static string ComputeSha256(string text)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = Encoding.UTF8.GetBytes(text ?? "");
                return BitConverter.ToString(sha.ComputeHash(bytes))
                    .Replace("-", "")
                    .ToLowerInvariant();
            }
        }

        private TextActionHistoryDocument LoadUnsafe()
        {
            try
            {
                if (!File.Exists(Path)) return EmptyDocument();
                var parsed = JsonConvert.DeserializeObject<TextActionHistoryDocument>(
                    File.ReadAllText(Path));
                if (parsed == null || parsed.Entries == null) return EmptyDocument();

                parsed.SchemaVersion = CurrentSchemaVersion;
                parsed.Entries = parsed.Entries
                    .Where(entry => entry != null && !string.IsNullOrWhiteSpace(entry.Id))
                    .TakeLastCompatible(MaxEntries)
                    .ToList();
                return parsed;
            }
            catch
            {
                return EmptyDocument();
            }
        }

        private void SaveUnsafe(TextActionHistoryDocument document)
        {
            document = document ?? EmptyDocument();
            document.SchemaVersion = CurrentSchemaVersion;
            document.Entries = (document.Entries ?? new List<TextActionHistoryEntry>())
                .Where(entry => entry != null)
                .TakeLastCompatible(MaxEntries)
                .ToList();

            var directory = System.IO.Path.GetDirectoryName(Path);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

            var json = JsonConvert.SerializeObject(
                document,
                Formatting.Indented,
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            var tempPath = Path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(tempPath, json, new UTF8Encoding(false));
                if (File.Exists(Path))
                {
                    File.Replace(tempPath, Path, null);
                }
                else
                {
                    File.Move(tempPath, Path);
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(tempPath)) File.Delete(tempPath);
                }
                catch
                {
                    // Best-effort cleanup must not mask the original I/O error.
                }
            }
        }

        private static TextActionHistoryDocument EmptyDocument()
        {
            return new TextActionHistoryDocument
            {
                SchemaVersion = CurrentSchemaVersion,
                Entries = new List<TextActionHistoryEntry>()
            };
        }

        private static string BuildPreview(string text)
        {
            text = text ?? "";
            var compact = new StringBuilder(text.Length);
            var pendingSpace = false;
            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                if (char.IsWhiteSpace(ch) || ch == '\a')
                {
                    pendingSpace = compact.Length > 0;
                    continue;
                }
                if (pendingSpace)
                {
                    compact.Append(' ');
                    pendingSpace = false;
                }
                compact.Append(ch);
            }

            var result = compact.ToString();
            if (result.Length <= PreviewMaxCharacters) return result;
            return result.Substring(0, PreviewMaxCharacters - 3) + "...";
        }
    }

    internal static class TextActionHistoryEnumerableExtensions
    {
        /// <summary>
        /// Enumerable.TakeLast is unavailable on .NET Framework 4.7.2.
        /// Materialize once and retain the tail without requiring newer BCLs.
        /// </summary>
        public static IEnumerable<T> TakeLastCompatible<T>(
            this IEnumerable<T> source,
            int count)
        {
            var list = (source ?? Enumerable.Empty<T>()).ToList();
            var skip = Math.Max(0, list.Count - Math.Max(0, count));
            return list.Skip(skip);
        }
    }
}
