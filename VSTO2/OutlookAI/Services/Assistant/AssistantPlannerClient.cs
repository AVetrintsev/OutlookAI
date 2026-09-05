using System;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Assistant
{
    public sealed class AssistantPlannerClient : IPlannerClient
    {
        private readonly LiteLlmChatService _chat;

        public AssistantPlannerClient(LiteLlmChatService chat)
        {
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
        }

        public async Task<TurnPlan> PlanAsync(
            TurnRequest request,
            TurnSnapshot snapshot,
            CancellationToken cancellationToken)
        {
            var response = await _chat.CompleteWithoutToolsAsync(
                BuildSystemPrompt(),
                BuildUserPrompt(request, snapshot),
                cancellationToken).ConfigureAwait(false);
            return ParsePlan(response) ?? FallbackPlan();
        }

        internal static TurnPlan ParsePlan(string response)
        {
            var obj = ExtractJsonObject(response);
            if (obj == null) return null;
            return new TurnPlan
            {
                Workflow = Normalize(ReadScalar(obj["workflow"]), "inbox_chat"),
                Intent = Normalize(ReadScalar(obj["intent"]), "unclear"),
                Confidence = NormalizeConfidence(ReadScalar(obj["confidence"])),
                ContextMode = NormalizeContextMode(ReadScalar(obj["context_mode"])),
                OutputContract = Normalize(ReadScalar(obj["output_contract"]), "plain_answer"),
                CapabilityIds = ((obj["capability_ids"] as JArray)?.Values<string>()
                    ?? Enumerable.Empty<string>())
                    .Where(id => !string.IsNullOrWhiteSpace(id))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray(),
                NeedsMoreData = (bool?)obj["needs_more_data"] ?? false,
                MissingData = ReadScalar(obj["missing_data"]).Trim(),
                Reason = ReadScalar(obj["reason"]).Trim()
            };
        }

        private static TurnPlan FallbackPlan()
        {
            return new TurnPlan
            {
                Workflow = "inbox_chat",
                Intent = "unclear",
                Confidence = "low",
                ContextMode = "metadata",
                OutputContract = "plain_answer",
                CapabilityIds = new string[0],
                Reason = "invalid_planner_json"
            };
        }

        private static JObject ExtractJsonObject(string response)
        {
            if (string.IsNullOrWhiteSpace(response)) return null;
            var start = response.IndexOf('{');
            var end = response.LastIndexOf('}');
            if (start < 0 || end <= start) return null;
            try { return JObject.Parse(response.Substring(start, end - start + 1)); }
            catch { return null; }
        }

        private static string BuildSystemPrompt()
        {
            return string.Join("\n", new[]
            {
                "You are the OutlookAI turn planner. Do not answer the user.",
                "Return only compact JSON.",
                "Schema: workflow, intent, confidence, context_mode, output_contract, capability_ids, needs_more_data, missing_data, reason.",
                "confidence: high | medium | low.",
                "context_mode: metadata | snippet | full | search_projection.",
                "Use metadata only when sender/recipient/date/attachment metadata is enough.",
                "Use snippet when the user asks about why, purpose, meaning, content, summary, or reply drafting for the selected item.",
                "Use search_projection for mailbox search/report questions.",
                "Pick capability_ids only from retrieved capabilities."
            });
        }

        private static string BuildUserPrompt(TurnRequest request, TurnSnapshot snapshot)
        {
            var sb = new StringBuilder();
            sb.AppendLine("User request:");
            sb.AppendLine(request?.UserText ?? "");
            sb.AppendLine();
            sb.AppendLine("Surface: " + (request?.Surface ?? "inbox_chat"));
            sb.AppendLine("Workflow: " + (request?.WorkflowId ?? "inbox_chat"));
            sb.AppendLine("Selected metadata:");
            var first = snapshot?.Selection?.Messages?.FirstOrDefault();
            if (first == null)
            {
                sb.AppendLine("none");
            }
            else
            {
                sb.AppendLine("id: " + (first.Id ?? ""));
                sb.AppendLine("item_type: " + (first.ItemType ?? "mail"));
                sb.AppendLine("direction: " + (first.Direction ?? "unknown"));
                sb.AppendLine("my_role: " + (first.MyRole ?? "unknown"));
                sb.AppendLine("subject: " + (first.Subject ?? ""));
                sb.AppendLine("from: " + (first.From ?? ""));
                sb.AppendLine("to: " + string.Join(", ", first.To ?? new string[0]));
                sb.AppendLine("is_from_me: " + first.IsFromMe);
                sb.AppendLine("is_to_me: " + first.IsToMe);
                sb.AppendLine("has_attachments: " + (first.Attachments != null && first.Attachments.Count > 0));
            }
            sb.AppendLine();
            sb.AppendLine("Retrieved capabilities:");
            foreach (var result in snapshot?.Capabilities ?? new CapabilitySearchResult[0])
            {
                var card = result.Card;
                if (card == null) continue;
                sb.AppendLine("- id: " + card.Id
                    + "; type: " + card.Type
                    + "; group: " + card.Group
                    + "; title: " + card.Title
                    + "; use_for: " + string.Join(", ", card.WhenToUse ?? new string[0])
                    + "; risk: " + card.Risk);
            }
            return sb.ToString();
        }

        private static string Normalize(string value, string fallback)
        {
            value = (value ?? "").Trim();
            return string.IsNullOrWhiteSpace(value) ? fallback : value;
        }

        private static string ReadScalar(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return "";
            }
            if (token.Type == JTokenType.String
                || token.Type == JTokenType.Integer
                || token.Type == JTokenType.Float
                || token.Type == JTokenType.Boolean)
            {
                return token.ToString();
            }
            return "";
        }

        private static string NormalizeConfidence(string value)
        {
            value = (value ?? "").Trim().ToLowerInvariant();
            return value == "high" || value == "medium" || value == "low" ? value : "low";
        }

        private static string NormalizeContextMode(string value)
        {
            value = (value ?? "").Trim().ToLowerInvariant();
            switch (value)
            {
                case "metadata":
                case "snippet":
                case "full":
                case "search_projection":
                    return value;
                default:
                    return "metadata";
            }
        }
    }
}
