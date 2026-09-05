using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Chat;

namespace OutlookAI.Services.Assistant
{
    public sealed class AssistantExecutorClient : IExecutorClient
    {
        private readonly LiteLlmChatService _chat;
        private readonly IToolHost _toolHost;

        public AssistantExecutorClient(LiteLlmChatService chat, IToolHost toolHost)
        {
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
        }

        public async Task<AssistantEngineResult> ExecuteAsync(
            TurnRequest request,
            ContextBundle context,
            ValidatedTurnPlan plan,
            ChatEventSink sink,
            CancellationToken cancellationToken)
        {
            var history = CompactHistory(request?.History);
            var conversation = new ConversationContext
            {
                SystemInstructions = BuildSystemInstructions(context, plan),
                EnableSkills = true,
                PinnedSkillIds = request?.PinnedSkillIds,
                History = history,
                IncludeWriteTools = plan.IncludeWriteTools,
                AllowedToolNames = plan.AllowedToolNames,
                ReasoningEffortOverride = request?.ReasoningEffortOverride
            };

            var turn = await _chat.RunTurnAsync(
                conversation,
                request?.UserText ?? "",
                _toolHost,
                sink ?? new ChatEventSink(),
                cancellationToken).ConfigureAwait(false);
            var normalized = NormalizeAssistantText(turn.FinalAssistantText);
            turn.FinalAssistantText = normalized;
            NormalizeLastAssistantMessage(turn, normalized);

            return new AssistantEngineResult
            {
                TurnResult = turn,
                Envelope = new ResponseEnvelope
                {
                    Kind = turn.StopReason == StopReason.Completed ? "final" : "partial",
                    AnswerMarkdown = normalized,
                    Confidence = plan.Plan?.Confidence ?? "medium"
                }
            };
        }

        private static List<JObject> CompactHistory(IList<JObject> history)
        {
            var items = (history ?? new List<JObject>())
                .Where(item => item != null && (string)item["type"] == "message")
                .Reverse()
                .Take(6)
                .Reverse()
                .Select(item => (JObject)item.DeepClone())
                .ToList();
            return items;
        }

        private static void NormalizeLastAssistantMessage(TurnResult turn, string normalized)
        {
            var appended = turn?.AppendedItems;
            if (appended == null || string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }
            var last = appended.LastOrDefault(item =>
                string.Equals((string)item["type"], "message", StringComparison.OrdinalIgnoreCase)
                && string.Equals((string)item["role"], "assistant", StringComparison.OrdinalIgnoreCase));
            if (last != null)
            {
                last["content"] = normalized;
            }
        }

        private static string BuildSystemInstructions(ContextBundle context, ValidatedTurnPlan plan)
        {
            return string.Join("\n", new[]
            {
                "You are OutlookAI Assistant inside Microsoft Outlook.",
                "Answer the user directly and concisely in the user's language.",
                "Do not expose raw JSON, tool results, hidden plan data, or implementation details.",
                "Use only the provided context bundle and, when tools are available, call only necessary tools.",
                "If the context is insufficient, say exactly what data is missing.",
                "Output plain Markdown, not a JSON envelope.",
                "",
                "Validated plan:",
                JsonConvert.SerializeObject(new
                {
                    workflow = plan.Plan?.Workflow ?? "inbox_chat",
                    intent = plan.Plan?.Intent ?? "unclear",
                    context_mode = plan.ContextMode,
                    output_contract = plan.Plan?.OutputContract ?? "plain_answer",
                    allowed_tools = plan.AllowedToolNames ?? new string[0]
                }, Formatting.None),
                "",
                "Context bundle:",
                (context?.Json ?? new JObject()).ToString(Formatting.None)
            });
        }

        internal static string NormalizeAssistantText(string text)
        {
            var normalized = (text ?? "").Trim();
            var prefixes = new[]
            {
                "### Assistant:",
                "## Assistant:",
                "# Assistant:",
                "Assistant:",
                "assistant:",
                "### Ответ:",
                "Ответ:"
            };
            foreach (var prefix in prefixes)
            {
                if (normalized.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return normalized.Substring(prefix.Length).TrimStart();
                }
            }
            return normalized;
        }
    }
}
