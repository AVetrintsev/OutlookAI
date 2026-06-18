using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OutlookAI.Services.Export;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// All Outlook OOM access used by Phase 2 tools flows through this interface.
    /// Production implementation (LiveOutlookSurface) marshals every call onto
    /// the Outlook STA UI thread via OutlookThreadMarshaller. Tests stub the
    /// methods they need without touching COM.
    /// </summary>
    public interface IOutlookSurface
    {
        ComposeStateResult GetCurrentComposeState(bool includeFullBody);
        IReadOnlyList<FolderResult> ListFolders();
        SearchResult SearchMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken));
        MessageDetail ReadMessage(string messageId, bool includeFullBody);
        IReadOnlyList<MessageDetail> ReadMessages(string[] ids, bool includeBody, int maxItems, CancellationToken ct = default(CancellationToken));
        int CountMessages(SearchMessagesArgs args, CancellationToken ct = default(CancellationToken));
        IReadOnlyList<AggregationBucket> AggregateMessages(AggregateMessagesArgs args, CancellationToken ct = default(CancellationToken));
        IReadOnlyList<ThreadSummary> ListRecentThreadsWith(string recipientEmail, int maxThreads);
        FileSavedResult ExportExcel(ExportExcelArgs args, CancellationToken ct = default(CancellationToken));
        FileSavedResult ExportPdf(ExportPdfArgs args, CancellationToken ct = default(CancellationToken));
        CreatedDraft CreateDraft(CreateDraftArgs args);
        CreatedDraft CreateReplyDraft(CreateReplyDraftArgs args);
        CreatedDraft CreateMeetingDraft(CreateMeetingDraftArgs args);
        void OpenItem(string itemId);
        void MarkAsRead(string messageId, bool read);
        void FlagMessage(string messageId, string flag);
        void SetCategory(string messageId, string category);
        /// <summary>
        /// Returns items currently selected in the active Explorer, or the
        /// current Inspector item when the surface was constructed without an
        /// Explorer reference.
        /// </summary>
        CurrentSelectionResult GetCurrentSelection(bool includeFullBodies, int maxItems);
    }

    public sealed class ComposeStateResult
    {
        public string Subject { get; set; }
        public IReadOnlyList<string> ToRecipients { get; set; }
        public IReadOnlyList<string> CcRecipients { get; set; }
        public IReadOnlyList<string> BccRecipients { get; set; }
        public string SenderName { get; set; }
        public string SenderEmail { get; set; }
        public MailboxIdentity CurrentUser { get; set; }
        public string ItemType { get; set; }
        public string Direction { get; set; }
        public string MyRole { get; set; }
        public string BodyPlaintext { get; set; }
        public bool BodyTruncated { get; set; }
        public InReplyTo InReplyTo { get; set; }
        public IReadOnlyList<AttachmentSummary> Attachments { get; set; }
    }

    public sealed class InReplyTo
    {
        public string ThreadTopic { get; set; }
        public IReadOnlyList<ThreadMessage> LastNMessages { get; set; }
    }

    public sealed class ThreadMessage
    {
        public string From { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public string Snippet { get; set; }
    }

    public sealed class AttachmentSummary
    {
        public string Filename { get; set; }
        public long SizeBytes { get; set; }
    }

    public sealed class FolderResult
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ParentId { get; set; }
        public int ItemCount { get; set; }
    }

    public sealed class SearchMessagesArgs
    {
        public string Query { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string SubjectContains { get; set; }
        public string BodyContains { get; set; }
        // Hidden backward-compat fields. These are accepted by the parser for
        // older conversation history but are no longer advertised to the model.
        public bool? HasAttachment { get; set; }
        public bool? IsUnread { get; set; }
        public bool? IsFlagged { get; set; }
        /// <summary>One of "low" | "normal" | "high"; null = unset.</summary>
        public string Importance { get; set; }
        public string FolderId { get; set; }
        public DateTimeOffset? DateFrom { get; set; }
        public DateTimeOffset? DateTo { get; set; }
        public int MaxResults { get; set; } = 25;
        /// <summary>"current_folder" | "all_mail" | "auto"; default auto.</summary>
        public string Scope { get; set; } = "auto";
        /// <summary>"newest" | "oldest"; default newest.</summary>
        public string SortOrder { get; set; } = "newest";
        /// <summary>"any" | "with" | "without"; default any.</summary>
        public string AttachmentFilter { get; set; } = "any";
        /// <summary>"any" | "read" | "unread"; default any.</summary>
        public string ReadStatus { get; set; } = "any";
        /// <summary>"any" | "flagged" | "unflagged"; default any.</summary>
        public string FlagStatus { get; set; } = "any";
        /// <summary>"any" | "low" | "normal" | "high"; default any.</summary>
        public string ImportanceFilter { get; set; } = "any";
    }

    public sealed class MessageSummary
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public string Snippet { get; set; }
        public bool HasAttachments { get; set; }
        /// <summary>
        /// Display name of the folder the message lives in. Populated by
        /// <see cref="SearchResultProjector"/> from the projection input; used
        /// by the bulk-export `folder` column. May be empty when the source
        /// path did not carry a folder name.
        /// </summary>
        public string FolderName { get; set; }
    }

    public sealed class MessageDetail
    {
        public string Id { get; set; }
        public string ItemType { get; set; }
        public string Direction { get; set; }
        public string MyRole { get; set; }
        public MailboxIdentity CurrentUser { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public IReadOnlyList<string> Cc { get; set; }
        public DateTimeOffset SentAt { get; set; }
        public DateTimeOffset ReceivedAt { get; set; }
        public bool IsFromMe { get; set; }
        public bool IsToMe { get; set; }
        public bool IsCcToMe { get; set; }
        public string BodyPlaintext { get; set; }
        public bool BodyTruncated { get; set; }
        public IReadOnlyList<AttachmentSummary> Attachments { get; set; }
        public string InReplyToMessageId { get; set; }
        public string ConversationTopic { get; set; }
        public string Organizer { get; set; }
        public IReadOnlyList<string> RequiredAttendees { get; set; }
        public IReadOnlyList<string> OptionalAttendees { get; set; }
        public DateTimeOffset? Start { get; set; }
        public DateTimeOffset? End { get; set; }
        public string Location { get; set; }
        public string MeetingState { get; set; }
    }

    public sealed class MailboxIdentity
    {
        public string DisplayName { get; set; }
        public string SmtpAddress { get; set; }
        public IReadOnlyList<string> Aliases { get; set; }
    }

    public static class OutlookContextClassifier
    {
        public static bool IdentityMatches(MailboxIdentity identity, string value)
        {
            var email = ExtractEmail(value);
            if (string.IsNullOrWhiteSpace(email) || identity == null) return false;
            if (EqualsEmail(identity.SmtpAddress, email)) return true;
            return (identity.Aliases ?? new string[0]).Any(alias => EqualsEmail(alias, email));
        }

        public static string MailDirection(MailboxIdentity identity, string from, IEnumerable<string> to, IEnumerable<string> cc)
        {
            if (IdentityMatches(identity, from)) return "outgoing";
            if ((to ?? new string[0]).Any(value => IdentityMatches(identity, value))
                || (cc ?? new string[0]).Any(value => IdentityMatches(identity, value)))
            {
                return "incoming";
            }
            return "unknown";
        }

        public static string MailRole(MailboxIdentity identity, string from, IEnumerable<string> to, IEnumerable<string> cc)
        {
            if (IdentityMatches(identity, from)) return "sender";
            if ((to ?? new string[0]).Any(value => IdentityMatches(identity, value))) return "recipient";
            if ((cc ?? new string[0]).Any(value => IdentityMatches(identity, value))) return "cc_recipient";
            return "unknown";
        }

        public static string MeetingDirection(MailboxIdentity identity, string organizer, IEnumerable<string> required, IEnumerable<string> optional)
        {
            if (IdentityMatches(identity, organizer)) return "outgoing";
            if ((required ?? new string[0]).Any(value => IdentityMatches(identity, value))
                || (optional ?? new string[0]).Any(value => IdentityMatches(identity, value)))
            {
                return "incoming";
            }
            return "unknown";
        }

        public static string MeetingRole(MailboxIdentity identity, string organizer, IEnumerable<string> required, IEnumerable<string> optional)
        {
            if (IdentityMatches(identity, organizer)) return "organizer";
            if ((required ?? new string[0]).Any(value => IdentityMatches(identity, value))) return "required_attendee";
            if ((optional ?? new string[0]).Any(value => IdentityMatches(identity, value))) return "optional_attendee";
            return "unknown";
        }

        public static string ExtractEmail(string value)
        {
            value = (value ?? "").Trim();
            if (value.Length == 0) return "";
            var lt = value.LastIndexOf('<');
            var gt = value.LastIndexOf('>');
            if (lt >= 0 && gt > lt) value = value.Substring(lt + 1, gt - lt - 1).Trim();
            return value.Trim('"').Trim();
        }

        private static bool EqualsEmail(string left, string right)
        {
            return string.Equals(ExtractEmail(left), ExtractEmail(right), StringComparison.OrdinalIgnoreCase);
        }
    }

    public sealed class ThreadSummary
    {
        public string ThreadTopic { get; set; }
        public DateTimeOffset LastMessageAt { get; set; }
        public int MessageCount { get; set; }
        public string Snippet { get; set; }
        public string ThreadId { get; set; }
    }

    public sealed class CreateDraftArgs
    {
        public string Subject { get; set; }
        public string BodyPlaintext { get; set; }
        public IReadOnlyList<string> To { get; set; }
        public IReadOnlyList<string> Cc { get; set; }
        public string InReplyToMessageId { get; set; }
    }

    public sealed class CreatedDraft
    {
        public string DraftId { get; set; }
        public string Location { get; set; }
        public string DisplayName { get; set; }
    }

    public sealed class CreateReplyDraftArgs
    {
        public string SourceMessageId { get; set; }
        public string BodyPlaintext { get; set; }
    }

    public sealed class CreateMeetingDraftArgs
    {
        public string SourceMessageId { get; set; }
        public string BodyPlaintext { get; set; }
    }

    /// <summary>
    /// Phase 3a: result of <see cref="IOutlookSurface.GetCurrentSelection"/>.
    /// Used by the Inbox Copilot's <c>outlook_get_current_selection</c> tool
    /// so the model can act on whichever message(s) the user has highlighted
    /// in the reading pane.
    /// </summary>
    public sealed class CurrentSelectionResult
    {
        /// <summary>Display name of the currently-active folder (e.g. "Inbox").</summary>
        public string Folder { get; set; }
        /// <summary>Stable short-id for the folder; matches outlook_list_folders ids.</summary>
        public string FolderId { get; set; }
        /// <summary>Total selection count (may exceed Messages.Count if MaxItems clamped).</summary>
        public int Count { get; set; }
        /// <summary>Up to MaxItems selected messages.</summary>
        public IReadOnlyList<MessageDetail> Messages { get; set; }
    }
}
