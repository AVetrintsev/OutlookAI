using System;
using System.Linq;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services.Assistant
{
    public sealed class AssistantContextResolver : IContextResolver
    {
        public ContextBundle Resolve(
            TurnRequest request,
            TurnSnapshot snapshot,
            ValidatedTurnPlan plan)
        {
            var root = new JObject(
                new JProperty("workflow", request?.WorkflowId ?? "inbox_chat"),
                new JProperty("surface", request?.Surface ?? "inbox_chat"),
                new JProperty("context_mode", plan?.ContextMode ?? "metadata"),
                new JProperty("folder", new JObject(
                    new JProperty("name", snapshot?.FolderName ?? "Inbox"),
                    new JProperty("unread", snapshot?.UnreadCount ?? 0),
                    new JProperty("total", snapshot?.TotalCount ?? 0))));

            var selected = snapshot?.Selection?.Messages?.FirstOrDefault();
            if (selected != null)
            {
                root["selected_item"] = BuildSelectedItem(selected, plan?.ContextMode ?? "metadata");
            }
            root["capabilities"] = new JArray((plan?.SelectedCapabilities ?? new CapabilitySearchResult[0])
                .Select(result => new JObject(
                    new JProperty("id", result.Card.Id),
                    new JProperty("group", result.Card.Group),
                    new JProperty("risk", result.Card.Risk),
                    new JProperty("tools", new JArray(result.Card.ToolNames ?? new string[0])))));

            return new ContextBundle
            {
                ContextMode = plan?.ContextMode ?? "metadata",
                Json = root
            };
        }

        private static JObject BuildSelectedItem(MessageDetail message, string contextMode)
        {
            var obj = new JObject(
                new JProperty("id", message.Id ?? ""),
                new JProperty("item_type", message.ItemType ?? "mail"),
                new JProperty("direction", message.Direction ?? "unknown"),
                new JProperty("my_role", message.MyRole ?? "unknown"),
                new JProperty("current_user", FormatIdentity(message.CurrentUser)),
                new JProperty("subject", message.Subject ?? ""),
                new JProperty("from", message.From ?? ""),
                new JProperty("to", new JArray(message.To ?? new string[0])),
                new JProperty("cc", new JArray(message.Cc ?? new string[0])),
                new JProperty("sent_at", message.SentAt.ToString("o")),
                new JProperty("received_at", message.ReceivedAt.ToString("o")),
                new JProperty("is_from_me", message.IsFromMe),
                new JProperty("is_to_me", message.IsToMe),
                new JProperty("is_cc_to_me", message.IsCcToMe),
                new JProperty("has_attachments", message.Attachments != null && message.Attachments.Count > 0));

            if (contextMode == "snippet" || contextMode == "full")
            {
                var max = contextMode == "full" ? 1600 : 500;
                var snippet = BuildSnippet(message.BodyPlaintext, max);
                if (!string.IsNullOrWhiteSpace(snippet))
                {
                    obj["body_snippet"] = snippet;
                }
            }
            return obj;
        }

        internal static string BuildSnippet(string body, int maxLength)
        {
            var snippet = (body ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            var markers = new[] { "-----Original Message-----", "From:", "От:", "вторник,", "понедельник,", "среда,", "четверг,", "пятница," };
            foreach (var marker in markers)
            {
                var index = snippet.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                if (index > 0)
                {
                    snippet = snippet.Substring(0, index).Trim();
                }
            }
            while (snippet.Contains("  "))
            {
                snippet = snippet.Replace("  ", " ");
            }
            return snippet.Length > maxLength ? snippet.Substring(0, maxLength) : snippet;
        }

        private static string FormatIdentity(MailboxIdentity identity)
        {
            if (identity == null) return "";
            if (!string.IsNullOrWhiteSpace(identity.DisplayName) && !string.IsNullOrWhiteSpace(identity.SmtpAddress))
            {
                return identity.DisplayName + " <" + identity.SmtpAddress + ">";
            }
            return identity.SmtpAddress ?? identity.DisplayName ?? "";
        }
    }
}
