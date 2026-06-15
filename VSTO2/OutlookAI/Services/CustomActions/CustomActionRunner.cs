using System;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Export;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services.CustomActions
{
    public sealed class CustomActionRunner
    {
        private const string MissingContextMessage =
            "Не удалось получить текст письма. Откройте или выберите сообщение и повторите действие.";

        private readonly LiteLlmChatService _chat;
        private readonly IOutlookSurface _surface;
        private readonly IToolHost _toolHost;
        private readonly CustomActionStateStore _state;

        public CustomActionRunner(
            LiteLlmChatService chat,
            IOutlookSurface surface,
            IToolHost toolHost = null,
            CustomActionStateStore state = null)
        {
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _surface = surface ?? throw new ArgumentNullException(nameof(surface));
            _toolHost = toolHost;
            _state = state ?? new CustomActionStateStore();
        }

        public async Task<CustomActionRunResult> RunAsync(
            CustomActionDefinition action,
            CancellationToken ct,
            ChatEventSink sink = null)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            sink = sink ?? new ChatEventSink();

            var context = BuildContext(action, ct);
            if (IsMissingContext(context))
            {
                return new CustomActionRunResult
                {
                    Text = MissingContextMessage,
                    Output = "chat"
                };
            }

            var userMessage = BuildControlledUserMessage(action, context);
            var allowedTools = (action.AllowedTools ?? new string[0])
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            string text;
            if (allowedTools.Length > 0 && _toolHost != null)
            {
                var turn = await _chat.RunTurnAsync(
                    new ConversationContext
                    {
                        SystemInstructions = PromptCatalog.Default.Get("custom_action_controlled"),
                        IncludeWriteTools = allowedTools.Any(ToolManifestCatalog.IsWriteTool),
                        AllowedToolNames = allowedTools
                    },
                    userMessage,
                    _toolHost,
                    sink,
                    ct).ConfigureAwait(false);
                text = turn.FinalAssistantText ?? "";
            }
            else
            {
                text = await _chat.CompleteWithoutToolsAsync(
                    PromptCatalog.Default.Get("custom_action_controlled"),
                    userMessage,
                    sink,
                    ct).ConfigureAwait(false);
            }

            var result = ApplyOutput(action, text, ct);
            _state.SetLastRun(action.Id, DateTimeOffset.UtcNow);
            return result;
        }

        private string BuildContext(CustomActionDefinition action, CancellationToken ct)
        {
            var ctx = action.Context ?? new CustomActionContext();
            var source = Normalize(ctx.Source, "current_selection");
            if (source == "current_open_message")
            {
                var compose = _surface.GetCurrentComposeState(ctx.IncludeFullBodies);
                if (HasComposeContent(compose))
                {
                    return FormatCompose(compose, ctx.IncludeAttachments);
                }

                var selection = _surface.GetCurrentSelection(
                    ctx.IncludeFullBodies,
                    Clamp(ctx.MaxItems, 1, 100, 20));
                if (HasSelectionContent(selection))
                {
                    return FormatSelection(selection, ctx.IncludeAttachments);
                }

                return MissingContextMessage;
            }

            if (source == "related_thread")
            {
                var selection = _surface.GetCurrentSelection(
                    includeFullBodies: false,
                    maxItems: 1);
                var selected = selection?.Messages?.FirstOrDefault();
                var topic = selected?.ConversationTopic;
                if (string.IsNullOrWhiteSpace(topic)) topic = selected?.Subject;
                if (!string.IsNullOrWhiteSpace(topic))
                {
                    var threadSearch = _surface.SearchMessages(new SearchMessagesArgs
                    {
                        Scope = "current_folder",
                        SubjectContains = topic,
                        MaxResults = Clamp(ctx.MaxItems, 1, 100, 20)
                    }, ct);
                    var threadIds = threadSearch?.Messages?
                        .Select(message => message.Id)
                        .Where(id => !string.IsNullOrWhiteSpace(id))
                        .ToArray() ?? new string[0];
                    if (threadIds.Length > 0)
                    {
                        var threadDetails = _surface.ReadMessages(
                            threadIds,
                            includeBody: ctx.IncludeFullBodies,
                            maxItems: threadIds.Length,
                            ct: ct);
                        return FormatDetails(threadDetails, ctx.IncludeAttachments);
                    }
                }
                return FormatSelection(
                    _surface.GetCurrentSelection(ctx.IncludeFullBodies, Clamp(ctx.MaxItems, 1, 100, 20)),
                    ctx.IncludeAttachments);
            }

            if (source == "current_selection" || source == "selected_messages")
            {
                var selection = _surface.GetCurrentSelection(
                    ctx.IncludeFullBodies,
                    Clamp(ctx.MaxItems, 1, 100, 20));
                return FormatSelection(selection, ctx.IncludeAttachments);
            }

            var args = new SearchMessagesArgs
            {
                Scope = source == "all_folders" || Normalize(ctx.FolderScope, "") == "all_folders"
                    ? "all_mail"
                    : "current_folder",
                ReadStatus = MapReadFilter(ctx.ReadFilter),
                MaxResults = Clamp(ctx.MaxItems, 1, 100, 20)
            };
            ApplyTimeRange(args, action.Id, ctx);
            var search = _surface.SearchMessages(args, ct);
            if (search == null || search.Messages == null || search.Messages.Count == 0)
            {
                return "Нет сообщений, соответствующих параметрам действия.";
            }

            if (!ctx.IncludeFullBodies)
            {
                return FormatSummaries(search.Messages);
            }

            var ids = search.Messages.Select(m => m.Id).Where(id => !string.IsNullOrWhiteSpace(id)).ToArray();
            var details = _surface.ReadMessages(ids, includeBody: true, maxItems: args.MaxResults, ct: ct);
            return FormatDetails(details, ctx.IncludeAttachments);
        }

        private CustomActionRunResult ApplyOutput(
            CustomActionDefinition action,
            string text,
            CancellationToken ct)
        {
            var output = Normalize(action.Output, "chat");
            if (output == "create_draft")
            {
                var draft = _surface.CreateDraft(new CreateDraftArgs
                {
                    Subject = action.Title ?? "OutlookAI",
                    BodyPlaintext = text ?? ""
                });
                return new CustomActionRunResult
                {
                    Text = text,
                    Output = output,
                    DraftId = draft?.DraftId
                };
            }
            if (output == "export_pdf")
            {
                var saved = _surface.ExportPdf(new ExportPdfArgs
                {
                    Title = action.Title,
                    ContentMarkdown = text ?? "",
                    FilenameHint = action.Id
                }, ct);
                return new CustomActionRunResult
                {
                    Text = text,
                    Output = output,
                    FilePath = saved?.Path
                };
            }
            if (output == "export_excel")
            {
                var saved = _surface.ExportExcel(new ExportExcelArgs
                {
                    FilenameHint = action.Id,
                    SheetName = "OutlookAI",
                    Columns = new[]
                    {
                        new ExcelColumnSpec { Name = "Action", Type = ExcelColumnType.Text },
                        new ExcelColumnSpec { Name = "Result", Type = ExcelColumnType.Text }
                    },
                    Rows = new[]
                    {
                        new JToken[] { action.Title ?? action.Id ?? "", text ?? "" }
                    }
                }, ct);
                return new CustomActionRunResult
                {
                    Text = text,
                    Output = output,
                    FilePath = saved?.Path
                };
            }

            return new CustomActionRunResult
            {
                Text = text,
                Output = output
            };
        }

        private string BuildControlledUserMessage(CustomActionDefinition action, string context)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Пользовательская инструкция:");
            sb.AppendLine(action.Prompt ?? "");
            sb.AppendLine();
            sb.AppendLine("Контекст:");
            sb.AppendLine(context ?? "");
            return sb.ToString();
        }

        private void ApplyTimeRange(SearchMessagesArgs args, string actionId, CustomActionContext ctx)
        {
            var now = DateTimeOffset.UtcNow;
            var range = Normalize(ctx.TimeRange, "today");
            if (range == "last_hour")
            {
                args.DateFrom = now.AddHours(-1);
            }
            else if (range == "today")
            {
                args.DateFrom = new DateTimeOffset(now.UtcDateTime.Date, TimeSpan.Zero);
            }
            else if (range == "yesterday")
            {
                var today = new DateTimeOffset(now.UtcDateTime.Date, TimeSpan.Zero);
                args.DateFrom = today.AddDays(-1);
                args.DateTo = today;
            }
            else if (range == "since_last_run")
            {
                args.DateFrom = _state.GetLastRun(actionId) ?? now.AddDays(-1);
            }
            else if (range == "manual")
            {
                args.DateFrom = ctx.ManualFrom;
                args.DateTo = ctx.ManualTo;
            }
        }

        private static string MapReadFilter(string value)
        {
            var normalized = Normalize(value, "any");
            if (normalized == "unread") return "unread";
            if (normalized == "read") return "read";
            return "any";
        }

        private static string FormatCompose(ComposeStateResult state, bool includeAttachments)
        {
            if (state == null) return "Открытое сообщение недоступно.";
            var sb = new StringBuilder();
            sb.AppendLine("Текущее открытое сообщение/черновик:");
            sb.AppendLine("Subject: " + (state.Subject ?? ""));
            sb.AppendLine("To: " + string.Join(", ", state.ToRecipients ?? new string[0]));
            sb.AppendLine("Cc: " + string.Join(", ", state.CcRecipients ?? new string[0]));
            if (includeAttachments && state.Attachments != null && state.Attachments.Count > 0)
            {
                sb.AppendLine("Attachments: " + string.Join(", ", state.Attachments.Select(a => a.Filename)));
            }
            sb.AppendLine("Body:");
            sb.AppendLine(state.BodyPlaintext ?? "");
            return sb.ToString();
        }

        private static string FormatSelection(CurrentSelectionResult selection, bool includeAttachments)
        {
            if (selection == null || selection.Messages == null || selection.Messages.Count == 0)
            {
                return "Нет выбранных сообщений.";
            }
            return FormatDetails(selection.Messages, includeAttachments);
        }

        private static bool HasComposeContent(ComposeStateResult state)
        {
            if (state == null) return false;
            return !string.IsNullOrWhiteSpace(state.Subject)
                || !string.IsNullOrWhiteSpace(state.BodyPlaintext)
                || (state.ToRecipients != null && state.ToRecipients.Any(r => !string.IsNullOrWhiteSpace(r)))
                || (state.CcRecipients != null && state.CcRecipients.Any(r => !string.IsNullOrWhiteSpace(r)));
        }

        private static bool HasSelectionContent(CurrentSelectionResult selection)
        {
            return selection != null
                && selection.Messages != null
                && selection.Messages.Any(m =>
                    m != null
                    && (!string.IsNullOrWhiteSpace(m.Subject)
                        || !string.IsNullOrWhiteSpace(m.BodyPlaintext)
                        || !string.IsNullOrWhiteSpace(m.From)));
        }

        private static bool IsMissingContext(string context)
        {
            return string.IsNullOrWhiteSpace(context)
                || string.Equals(context.Trim(), MissingContextMessage, StringComparison.Ordinal);
        }

        private static string FormatSummaries(System.Collections.Generic.IEnumerable<MessageSummary> messages)
        {
            var sb = new StringBuilder();
            foreach (var m in messages)
            {
                sb.AppendLine("- " + (m.Subject ?? ""));
                sb.AppendLine("  From: " + (m.From ?? ""));
                sb.AppendLine("  Received: " + m.ReceivedAt.ToString("o"));
                sb.AppendLine("  Snippet: " + (m.Snippet ?? ""));
            }
            return sb.ToString();
        }

        private static string FormatDetails(System.Collections.Generic.IEnumerable<MessageDetail> messages, bool includeAttachments)
        {
            var sb = new StringBuilder();
            foreach (var m in messages ?? Enumerable.Empty<MessageDetail>())
            {
                sb.AppendLine("---");
                sb.AppendLine("Subject: " + (m.Subject ?? ""));
                sb.AppendLine("From: " + (m.From ?? ""));
                sb.AppendLine("To: " + string.Join(", ", m.To ?? new string[0]));
                sb.AppendLine("Received: " + m.ReceivedAt.ToString("o"));
                if (includeAttachments && m.Attachments != null && m.Attachments.Count > 0)
                {
                    sb.AppendLine("Attachments: " + string.Join(", ", m.Attachments.Select(a => a.Filename)));
                }
                sb.AppendLine("Body:");
                sb.AppendLine(m.BodyPlaintext ?? "");
            }
            return sb.ToString();
        }

        private static int Clamp(int value, int min, int max, int fallback)
        {
            if (value <= 0) value = fallback;
            return Math.Max(min, Math.Min(max, value));
        }

        private static string Normalize(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value)
                ? fallback
                : value.Trim().ToLowerInvariant();
        }
    }
}
