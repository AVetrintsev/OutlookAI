using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    public interface IConversationSurface
    {
        ConversationResult ReadConversation(string messageId, int maxItems, CancellationToken ct);
    }

    public sealed class ConversationResult
    {
        public string ConversationId { get; set; }
        public IReadOnlyList<MessageDetail> Messages { get; set; } = new MessageDetail[0];
        public bool Truncated { get; set; }
        public int UnavailableCount { get; set; }
        public string CoverageNote { get; set; }

        public string CoverageSummary => "Проанализировано сообщений: " + (Messages?.Count ?? 0)
            + (Truncated ? ". Цепочка получена не полностью" : "")
            + (UnavailableCount > 0 ? ". Недоступно: " + UnavailableCount : "")
            + ((Messages ?? new MessageDetail[0]).Any(message => message.BodyTruncated) ? ". Часть текстов сокращена или недоступна" : "")
            + ". Содержимое вложений не анализировалось. " + (CoverageNote ?? "");

        public JObject ToJson(bool includeBodies = true, bool includeAttachmentMetadata = true)
        {
            return new JObject(
                new JProperty("conversation_id", ConversationId ?? ""),
                new JProperty("messages", new JArray(Messages.OrderBy(MessageDate).Select(message =>
                {
                    var detail = OutlookJsonProjection.MessageDetail(message, includeBodies);
                    if (!includeAttachmentMetadata) detail.Remove("attachments");
                    return detail;
                }))),
                new JProperty("message_bodies_included", includeBodies),
                new JProperty("attachment_metadata_included", includeAttachmentMetadata),
                new JProperty("truncated", Truncated),
                new JProperty("unavailable_count", UnavailableCount),
                new JProperty("coverage_note", CoverageNote ?? ""),
                new JProperty("coverage_summary", CoverageSummary),
                new JProperty("attachment_contents_analyzed", false));
        }

        public static DateTimeOffset MessageDate(MessageDetail message) => message.SentAt > DateTimeOffset.MinValue
            ? message.SentAt : message.ReceivedAt;

        public static ConversationResult ReadCurrent(IOutlookSurface surface, int maxItems, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var selected = surface?.GetCurrentSelection(true, 1)?.Messages?.FirstOrDefault();
            if (selected == null) throw new InvalidOperationException("Выберите или откройте письмо для анализа переписки.");
            if (surface is IConversationSurface conversations)
                return conversations.ReadConversation(selected.Id, maxItems, ct);
            return new ConversationResult { Messages = new[] { selected }, Truncated = true,
                CoverageNote = "Доступно только выбранное письмо; связанная переписка не получена." };
        }
    }

    public sealed class OutlookReadConversationTool : IOutlookTool
    {
        public string Name => "outlook_read_conversation";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            var id = (string)args["message_id"];
            var maxItems = Math.Max(1, Math.Min(100, (int?)args["max_items"] ?? 100));
            ConversationResult result;
            if (string.IsNullOrWhiteSpace(id)) result = ConversationResult.ReadCurrent(surface, maxItems, ct);
            else if (surface is IConversationSurface conversations) result = conversations.ReadConversation(id, maxItems, ct);
            else result = new ConversationResult { Truncated = true, CoverageNote = "Связанная переписка недоступна." };
            return Task.FromResult(result.ToJson().ToString(Newtonsoft.Json.Formatting.None));
        }
    }
}
