using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.Services.Tools
{
    public sealed partial class LiveOutlookSurface
    {
        public ConversationResult ReadConversation(string messageId, int maxItems, CancellationToken ct)
        {
            maxItems = Math.Max(1, Math.Min(100, maxItems));
            return _marshaller.RunAsync(() => ReadConversationCore(messageId, maxItems, ct), ct).GetAwaiter().GetResult();
        }

        private ConversationResult ReadConversationCore(string messageId, int maxItems, CancellationToken ct)
        {
            const string storeEntryId = "http://schemas.microsoft.com/mapi/proptag/0x0FFB0102";
            object seed = null;
            Outlook.Conversation conversation = null;
            Outlook.Table table = null;
            var result = new ConversationResult
            {
                CoverageNote = "Переписка Outlook во всех доступных папках и хранилищах. Удалённые элементы автоматически не подбираются."
            };
            var messages = new List<MessageDetail>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                ct.ThrowIfCancellationRequested();
                seed = ResolveConversationSeed(messageId, ct);
                if (seed == null) throw new InvalidOperationException("Исходное письмо недоступно.");
                var selected = BuildSelectionDetail(seed, true);
                if (selected != null) { messages.Add(selected); seen.Add(selected.Id ?? ""); }
                if (seed is Outlook.MailItem mail)
                {
                    result.ConversationId = mail.ConversationID;
                    conversation = mail.GetConversation();
                }
                else if (seed is Outlook.MeetingItem meeting)
                {
                    result.ConversationId = meeting.ConversationID;
                    conversation = meeting.GetConversation();
                }
                if (conversation == null)
                {
                    result.Truncated = true;
                    result.CoverageNote = "Outlook не предоставил цепочку. Анализируется только открытый элемент.";
                }
                else
                {
                    table = conversation.GetTable();
                    table.Columns.Add(storeEntryId);
                    // Sort is supported for explicit scalar Outlook properties; reading
                    // newest first keeps later confirmations when the thread exceeds the cap.
                    try { table.Sort("[SentOn]", true); }
                    catch (COMException)
                    {
                        table.Sort("[CreationTime]", true);
                        result.CoverageNote += " Дата отправки для отбора недоступна: при ограничении используется время создания элемента.";
                    }
                    result.Truncated = table.GetRowCount() > maxItems;
                    var scanned = 0;
                    while (!table.EndOfTable && scanned < 1000)
                    {
                        ct.ThrowIfCancellationRequested();
                        scanned++;
                        Outlook.Row row = null;
                        object item = null;
                        try
                        {
                            row = table.GetNextRow();
                            var entryId = Convert.ToString(row["EntryID"]);
                            if (string.IsNullOrWhiteSpace(entryId)) { result.UnavailableCount++; continue; }
                            var shortId = _ids.Shorten(entryId);
                            if (seen.Contains(shortId)) continue;
                            if (messages.Count >= maxItems) { result.Truncated = true; break; }
                            item = _application.Session.GetItemFromID(entryId, row.BinaryToString(storeEntryId));
                            var detail = BuildSelectionDetail(item, true);
                            if (detail == null) { result.UnavailableCount++; continue; }
                            seen.Add(shortId);
                            messages.Add(detail);
                        }
                        catch (COMException) { result.UnavailableCount++; }
                        finally { ReleaseConversationObject(item); ReleaseConversationObject(row); }
                        YieldUi(ct);
                    }
                    if (!table.EndOfTable) result.Truncated = true;
                }
            }
            catch (COMException)
            {
                result.Truncated = true;
                result.CoverageNote = "Часть переписки недоступна в Outlook. Выводы ограничены полученными письмами.";
                result.UnavailableCount++;
            }
            finally
            {
                ReleaseConversationObject(table);
                ReleaseConversationObject(conversation);
                ReleaseConversationObject(seed);
            }
            // Share a bounded character budget fairly so a long early message cannot
            // silently displace later confirmations of completion.
            var perMessageLimit = messages.Count == 0 ? MaxBodyChars : Math.Min(MaxBodyChars, 240000 / messages.Count);
            foreach (var message in messages)
            {
                if ((message.BodyPlaintext ?? "").Length <= perMessageLimit) continue;
                message.BodyPlaintext = message.BodyPlaintext.Substring(0, perMessageLimit);
                message.BodyTruncated = true;
            }
            result.Messages = messages.OrderBy(ConversationResult.MessageDate).ToArray();
            return result;
        }

        private static void ReleaseConversationObject(object value)
        {
            try { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
            catch (InvalidComObjectException) { }
        }

        private object ResolveConversationSeed(string messageId, CancellationToken ct)
        {
            var entryId = string.IsNullOrWhiteSpace(messageId) ? null : _ids.Resolve(messageId);
            if (_composeInspector != null)
            {
                var current = _composeInspector.CurrentItem;
                if (string.IsNullOrEmpty(entryId)) return current;
                var id = current is Outlook.MailItem mail ? mail.EntryID
                    : current is Outlook.MeetingItem meeting ? meeting.EntryID : null;
                if (string.Equals(id, entryId, StringComparison.OrdinalIgnoreCase)) return current;
                ReleaseConversationObject(current);
            }
            if (string.IsNullOrEmpty(entryId)) return null;
            try { return _application.Session.GetItemFromID(entryId); }
            catch (COMException) { }
            // A shortened ID identifies the item, not its store. Resolve explicit IDs
            // from shared mailboxes/archives without assuming the default store.
            Outlook.Stores stores = _application.Session.Stores;
            try
            {
                for (var i = 1; i <= stores.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    Outlook.Store store = null;
                    try { store = stores[i]; return _application.Session.GetItemFromID(entryId, store.StoreID); }
                    catch (COMException) { }
                    finally { ReleaseConversationObject(store); }
                }
            }
            finally { ReleaseConversationObject(stores); }
            return null;
        }
    }
}
