using System.Linq;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    internal static class OutlookJsonProjection
    {
        public static JObject ProjectMailboxIdentity(MailboxIdentity identity)
        {
            identity = identity ?? new MailboxIdentity();
            return new JObject(
                new JProperty("display_name", identity.DisplayName ?? ""),
                new JProperty("smtp_address", identity.SmtpAddress ?? ""),
                new JProperty("aliases", new JArray((identity.Aliases ?? new string[0]).Cast<object>())));
        }

        public static JObject MessageDetail(MessageDetail detail, bool includeBody)
        {
            detail = detail ?? new MessageDetail();
            var json = new JObject(
                new JProperty("id", detail.Id ?? ""),
                new JProperty("item_type", detail.ItemType ?? "mail"),
                new JProperty("direction", detail.Direction ?? "unknown"),
                new JProperty("my_role", detail.MyRole ?? "unknown"),
                new JProperty("current_user", ProjectMailboxIdentity(detail.CurrentUser)),
                new JProperty("subject", detail.Subject ?? ""),
                new JProperty("from", detail.From ?? ""),
                new JProperty("to", new JArray((detail.To ?? new string[0]).Cast<object>())),
                new JProperty("cc", new JArray((detail.Cc ?? new string[0]).Cast<object>())),
                new JProperty("sent_at", detail.SentAt.ToString("o")),
                new JProperty("received_at", detail.ReceivedAt.ToString("o")),
                new JProperty("is_from_me", detail.IsFromMe),
                new JProperty("is_to_me", detail.IsToMe),
                new JProperty("is_cc_to_me", detail.IsCcToMe),
                new JProperty("body_truncated", detail.BodyTruncated),
                new JProperty("attachments", new JArray(
                    (detail.Attachments ?? new AttachmentSummary[0]).Select(a =>
                        new JObject(
                            new JProperty("filename", a.Filename ?? ""),
                            new JProperty("size_bytes", a.SizeBytes))))),
                new JProperty("in_reply_to_message_id", detail.InReplyToMessageId ?? ""),
                new JProperty("conversation_topic", detail.ConversationTopic ?? ""));

            if (includeBody)
            {
                json.Add("body_plaintext", detail.BodyPlaintext ?? "");
            }
            if ((detail.ItemType ?? "") == "meeting"
                || !string.IsNullOrWhiteSpace(detail.Organizer)
                || !string.IsNullOrWhiteSpace(detail.MeetingState))
            {
                json.Add("organizer", detail.Organizer ?? "");
                json.Add("required_attendees", new JArray((detail.RequiredAttendees ?? new string[0]).Cast<object>()));
                json.Add("optional_attendees", new JArray((detail.OptionalAttendees ?? new string[0]).Cast<object>()));
                json.Add("start", detail.Start.HasValue ? detail.Start.Value.ToString("o") : "");
                json.Add("end", detail.End.HasValue ? detail.End.Value.ToString("o") : "");
                json.Add("location", detail.Location ?? "");
                json.Add("meeting_state", detail.MeetingState ?? "");
            }

            return json;
        }

        public static void AddComposeContext(JObject json, ComposeStateResult state)
        {
            if (json == null || state == null) return;
            json.Add("item_type", state.ItemType ?? "mail");
            json.Add("direction", state.Direction ?? "unknown");
            json.Add("my_role", state.MyRole ?? "unknown");
            json.Add("current_user", ProjectMailboxIdentity(state.CurrentUser));
        }
    }
}
