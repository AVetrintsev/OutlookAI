using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services.CustomActions
{
    public sealed class CustomActionRecommendationService
    {
        private readonly LiteLlmChatService _chat;

        public CustomActionRecommendationService(LiteLlmChatService chat)
        {
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
        }

        public async Task<string[]> RecommendAsync(
            CurrentSelectionResult selection,
            IEnumerable<CustomActionGroup> groups,
            CancellationToken cancellationToken)
        {
            var available = (groups ?? Enumerable.Empty<CustomActionGroup>())
                .SelectMany(group => group.Actions ?? new CustomActionDefinition[0])
                .Where(action => !action.Disabled
                    && !string.IsNullOrWhiteSpace(action.Id)
                    && CustomActionSurface.Normalize(action.Surface) == CustomActionSurface.Assistant)
                .GroupBy(action => action.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            if (available.Count == 0 || selection?.Messages == null || selection.Messages.Count == 0)
            {
                return new string[0];
            }

            var prompt = new StringBuilder();
            prompt.AppendLine("Письмо:");
            foreach (var message in selection.Messages.Take(3))
            {
                prompt.AppendLine("ID: " + message.Id);
                prompt.AppendLine("Тема: " + (message.Subject ?? ""));
                prompt.AppendLine("От: " + (message.From ?? ""));
                prompt.AppendLine("Текст: " + Clip(message.BodyPlaintext, 1200));
            }
            prompt.AppendLine();
            prompt.AppendLine("Доступные действия:");
            foreach (var action in available.Values)
            {
                prompt.AppendLine(action.Id + " | " + action.Title + " | " + (action.Description ?? ""));
            }

            var response = await _chat.CompleteWithoutToolsAsync(
                "Выбери до 5 наиболее полезных действий для письма. Ответь только JSON: {\"action_ids\":[\"id\"]}. Не придумывай ID.",
                prompt.ToString(),
                cancellationToken).ConfigureAwait(false);
            var json = ExtractJson(response);
            var ids = (json?["action_ids"] as JArray)?.Values<string>()
                ?? Enumerable.Empty<string>();
            return ids
                .Where(id => !string.IsNullOrWhiteSpace(id) && available.ContainsKey(id))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(5)
                .ToArray();
        }

        public static string CacheKey(CurrentSelectionResult selection)
        {
            if (selection?.Messages == null || selection.Messages.Count == 0) return "";
            return string.Join("|", selection.Messages.Select(message => message.Id ?? "").Where(id => id.Length > 0));
        }

        private static JObject ExtractJson(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            var start = text.IndexOf('{');
            var end = text.LastIndexOf('}');
            if (start < 0 || end <= start) return null;
            try { return JObject.Parse(text.Substring(start, end - start + 1)); }
            catch { return null; }
        }

        private static string Clip(string value, int max)
        {
            value = value ?? "";
            return value.Length <= max ? value : value.Substring(0, max);
        }
    }
}
