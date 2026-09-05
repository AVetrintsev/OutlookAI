using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Services.Assistant
{
    public sealed class LocalHashEmbeddingProvider : IEmbeddingProvider
    {
        public int Dimension { get; private set; }

        public LocalHashEmbeddingProvider(int dimension = 256)
        {
            Dimension = Math.Max(32, dimension);
        }

        public Task<float[]> EmbedAsync(string text, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var vector = new float[Dimension];
            foreach (var token in TextTokens.Tokenize(text))
            {
                var hash = StableHash(token);
                var index = (int)(hash % (uint)Dimension);
                vector[index] += 1f;
            }
            Normalize(vector);
            return Task.FromResult(vector);
        }

        internal static void Normalize(float[] vector)
        {
            var length = Math.Sqrt(vector.Sum(v => (double)v * v));
            if (length <= 0) return;
            for (int i = 0; i < vector.Length; i++)
            {
                vector[i] = (float)(vector[i] / length);
            }
        }

        private static uint StableHash(string value)
        {
            unchecked
            {
                const uint fnvOffset = 2166136261;
                const uint fnvPrime = 16777619;
                var hash = fnvOffset;
                foreach (var ch in value ?? "")
                {
                    hash ^= ch;
                    hash *= fnvPrime;
                }
                return hash;
            }
        }
    }
}
