using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OutlookAI.Diagnostics;

namespace OutlookAI.Services.Assistant
{
    public sealed class LocalCapabilityIndex : ICapabilityIndex
    {
        private readonly CapabilityCardSource _source;
        private readonly IEmbeddingProvider _embeddings;
        private readonly bool _includeWriteTools;
        private readonly string _path;
        private List<Entry> _entries;

        public LocalCapabilityIndex(
            CapabilityCardSource source,
            IEmbeddingProvider embeddings,
            bool includeWriteTools,
            string path = null)
        {
            _source = source ?? throw new ArgumentNullException(nameof(source));
            _embeddings = embeddings ?? throw new ArgumentNullException(nameof(embeddings));
            _includeWriteTools = includeWriteTools;
            _path = path ?? DefaultPath();
        }

        public async Task<CapabilitySearchResult[]> SearchAsync(
            string query,
            TurnSnapshot snapshot,
            int topK,
            CancellationToken cancellationToken)
        {
            await EnsureLoadedAsync(cancellationToken).ConfigureAwait(false);
            var queryText = BuildQuery(query, snapshot);
            var queryVector = await _embeddings.EmbedAsync(queryText, cancellationToken).ConfigureAwait(false);
            var queryTokens = new HashSet<string>(TextTokens.Tokenize(queryText), StringComparer.OrdinalIgnoreCase);
            return _entries
                .Select(entry => Score(entry, queryVector, queryTokens))
                .Where(result => result.Score > 0)
                .OrderByDescending(result => result.Score)
                .Take(Math.Max(1, topK))
                .ToArray();
        }

        private async Task EnsureLoadedAsync(CancellationToken cancellationToken)
        {
            if (_entries != null) return;

            var cards = _source.LoadCards(_includeWriteTools);
            var sourceHash = HashCards(cards);
            var persisted = TryReadPersisted(sourceHash);
            if (persisted != null)
            {
                _entries = persisted;
                return;
            }

            var built = new List<Entry>();
            foreach (var card in cards)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var text = CardText(card);
                built.Add(new Entry
                {
                    SourceHash = sourceHash,
                    Card = card,
                    Text = text,
                    Tokens = TextTokens.Tokenize(text).ToArray(),
                    Vector = await _embeddings.EmbedAsync(text, cancellationToken).ConfigureAwait(false)
                });
            }
            _entries = built;
            TryWritePersisted(sourceHash, built);
        }

        private List<Entry> TryReadPersisted(string sourceHash)
        {
            try
            {
                if (!File.Exists(_path)) return null;
                var file = JsonConvert.DeserializeObject<IndexFile>(File.ReadAllText(_path));
                if (file == null || file.SourceHash != sourceHash || file.Dimension != _embeddings.Dimension)
                {
                    return null;
                }
                return file.Entries ?? new List<Entry>();
            }
            catch (Exception ex)
            {
                TraceLog.Write("Capability index read fallback: " + ex.Message, "AssistantEngine");
                return null;
            }
        }

        private void TryWritePersisted(string sourceHash, List<Entry> entries)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_path));
                File.WriteAllText(_path, JsonConvert.SerializeObject(new IndexFile
                {
                    SourceHash = sourceHash,
                    Dimension = _embeddings.Dimension,
                    CreatedAt = DateTimeOffset.UtcNow,
                    Entries = entries
                }, Formatting.None));
            }
            catch (Exception ex)
            {
                TraceLog.Write("Capability index write skipped: " + ex.Message, "AssistantEngine");
            }
        }

        private static CapabilitySearchResult Score(Entry entry, float[] queryVector, HashSet<string> queryTokens)
        {
            var vector = Cosine(queryVector, entry.Vector);
            var lexical = LexicalScore(entry.Tokens, queryTokens);
            var score = vector + lexical;
            return new CapabilitySearchResult
            {
                Card = entry.Card,
                Score = score,
                Reason = "vector=" + vector.ToString("0.000") + " lexical=" + lexical.ToString("0.000")
            };
        }

        private static double Cosine(float[] a, float[] b)
        {
            if (a == null || b == null || a.Length == 0 || a.Length != b.Length) return 0;
            double dot = 0;
            for (int i = 0; i < a.Length; i++)
            {
                dot += a[i] * b[i];
            }
            return dot;
        }

        private static double LexicalScore(string[] tokens, HashSet<string> queryTokens)
        {
            if (tokens == null || queryTokens == null || queryTokens.Count == 0) return 0;
            var matches = tokens.Count(token => queryTokens.Contains(token));
            return Math.Min(1.0, matches / 4.0);
        }

        private static string CardText(CapabilityCard card)
        {
            return string.Join(" ", new[]
            {
                card.Id,
                card.Type,
                card.Group,
                card.Title,
                card.Description,
                string.Join(" ", card.WhenToUse ?? new string[0]),
                string.Join(" ", card.WhenNotToUse ?? new string[0]),
                string.Join(" ", card.ToolNames ?? new string[0])
            });
        }

        private static string BuildQuery(string query, TurnSnapshot snapshot)
        {
            var selected = snapshot?.Selection?.Messages?.FirstOrDefault();
            return string.Join(" ", new[]
            {
                query ?? "",
                snapshot?.Surface ?? "",
                snapshot?.WorkflowId ?? "",
                selected?.ItemType ?? "",
                selected?.Direction ?? "",
                selected?.Subject ?? "",
                selected?.From ?? ""
            });
        }

        private static string HashCards(IEnumerable<CapabilityCard> cards)
        {
            var text = JsonConvert.SerializeObject(cards.OrderBy(card => card.Id), Formatting.None);
            using (var sha = SHA256.Create())
            {
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant();
            }
        }

        private static string DefaultPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI",
                "Indexes",
                "capabilities.index.json");
        }

        internal sealed class IndexFile
        {
            public string SourceHash { get; set; }
            public int Dimension { get; set; }
            public DateTimeOffset CreatedAt { get; set; }
            public List<Entry> Entries { get; set; }
        }

        internal sealed class Entry
        {
            public string SourceHash { get; set; }
            public CapabilityCard Card { get; set; }
            public string Text { get; set; }
            public string[] Tokens { get; set; }
            public float[] Vector { get; set; }
        }
    }
}
