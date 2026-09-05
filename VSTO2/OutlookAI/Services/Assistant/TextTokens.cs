using System.Collections.Generic;
using System.Text;

namespace OutlookAI.Services.Assistant
{
    internal static class TextTokens
    {
        public static IEnumerable<string> Tokenize(string text)
        {
            var token = new StringBuilder();
            foreach (var ch in (text ?? "").ToLowerInvariant())
            {
                if (char.IsLetterOrDigit(ch))
                {
                    token.Append(ch);
                    continue;
                }

                if (token.Length > 1)
                {
                    yield return token.ToString();
                }
                token.Clear();
            }
            if (token.Length > 1)
            {
                yield return token.ToString();
            }
        }
    }
}
