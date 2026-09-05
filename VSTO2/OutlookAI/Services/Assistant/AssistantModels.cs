using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services.Assistant
{
    public sealed class TurnRequest
    {
        public string Surface { get; set; } = "inbox_chat";
        public string WorkflowId { get; set; } = "inbox_chat";
        public string ConversationId { get; set; } = "";
        public string UserText { get; set; } = "";
        public string Locale { get; set; } = "ru-RU";
        public string FolderName { get; set; } = "Inbox";
        public int UnreadCount { get; set; }
        public int TotalCount { get; set; }
        public string ReasoningEffortOverride { get; set; }
        public string[] PinnedSkillIds { get; set; } = new string[0];
        public IList<JObject> History { get; set; } = new List<JObject>();
    }

    public sealed class TurnSnapshot
    {
        public string Surface { get; set; } = "inbox_chat";
        public string WorkflowId { get; set; } = "inbox_chat";
        public string FolderName { get; set; } = "Inbox";
        public int UnreadCount { get; set; }
        public int TotalCount { get; set; }
        public CurrentSelectionResult Selection { get; set; }
        public IReadOnlyList<CapabilitySearchResult> Capabilities { get; set; } =
            new List<CapabilitySearchResult>();
    }

    public sealed class TurnPlan
    {
        public string Workflow { get; set; } = "inbox_chat";
        public string Intent { get; set; } = "unclear";
        public string Confidence { get; set; } = "low";
        public string ContextMode { get; set; } = "metadata";
        public string OutputContract { get; set; } = "plain_answer";
        public string[] CapabilityIds { get; set; } = new string[0];
        public bool NeedsMoreData { get; set; }
        public string MissingData { get; set; } = "";
        public string Reason { get; set; } = "";
    }

    public sealed class ValidatedTurnPlan
    {
        public TurnPlan Plan { get; set; }
        public string ContextMode { get; set; } = "metadata";
        public string[] AllowedToolNames { get; set; } = new string[0];
        public bool IncludeWriteTools { get; set; }
        public bool NeedsMoreData { get; set; }
        public string MissingData { get; set; } = "";
        public IReadOnlyList<CapabilitySearchResult> SelectedCapabilities { get; set; } =
            new List<CapabilitySearchResult>();
    }

    public sealed class ContextBundle
    {
        public string ContextMode { get; set; } = "metadata";
        public JObject Json { get; set; } = new JObject();
    }

    public sealed class ResponseEnvelope
    {
        public string Kind { get; set; } = "final";
        public string AnswerMarkdown { get; set; } = "";
        public string[] Actions { get; set; } = new string[0];
        public string[] Artifacts { get; set; } = new string[0];
        public string MissingData { get; set; } = "";
        public string[] UsedRefs { get; set; } = new string[0];
        public string Confidence { get; set; } = "medium";
    }

    public sealed class AssistantEngineResult
    {
        public ResponseEnvelope Envelope { get; set; } = new ResponseEnvelope();
        public TurnResult TurnResult { get; set; } = new TurnResult();
    }
}
