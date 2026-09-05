using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services.Assistant
{
    public sealed class AssistantEngine : IAssistantEngine
    {
        private readonly IOutlookSurface _surface;
        private readonly ICapabilityIndex _capabilities;
        private readonly IPlannerClient _planner;
        private readonly IPlanValidator _validator;
        private readonly IContextResolver _contextResolver;
        private readonly IExecutorClient _executor;
        private readonly bool _includeWriteTools;

        public AssistantEngine(
            IOutlookSurface surface,
            ICapabilityIndex capabilities,
            IPlannerClient planner,
            IPlanValidator validator,
            IContextResolver contextResolver,
            IExecutorClient executor,
            bool includeWriteTools)
        {
            _surface = surface ?? throw new ArgumentNullException(nameof(surface));
            _capabilities = capabilities ?? throw new ArgumentNullException(nameof(capabilities));
            _planner = planner ?? throw new ArgumentNullException(nameof(planner));
            _validator = validator ?? throw new ArgumentNullException(nameof(validator));
            _contextResolver = contextResolver ?? throw new ArgumentNullException(nameof(contextResolver));
            _executor = executor ?? throw new ArgumentNullException(nameof(executor));
            _includeWriteTools = includeWriteTools;
        }

        public async Task<AssistantEngineResult> RunAsync(
            TurnRequest request,
            ChatEventSink sink,
            CancellationToken cancellationToken)
        {
            request = request ?? new TurnRequest();
            var snapshot = CaptureSnapshot(request, includeFullBodies: false);
            var retrieved = await _capabilities.SearchAsync(
                request.UserText ?? "",
                snapshot,
                8,
                cancellationToken).ConfigureAwait(false);
            snapshot.Capabilities = retrieved;

            var plan = await _planner.PlanAsync(request, snapshot, cancellationToken).ConfigureAwait(false);
            var validated = _validator.Validate(plan, request, snapshot, _includeWriteTools);
            TraceLog.Write(
                "intent=" + (plan?.Intent ?? "")
                + " context=" + validated.ContextMode
                + " tools=" + string.Join(",", validated.AllowedToolNames ?? new string[0])
                + " missing=" + validated.NeedsMoreData,
                "AssistantEngine");

            if (validated.NeedsMoreData)
            {
                var answer = string.IsNullOrWhiteSpace(validated.MissingData)
                    ? "Мне не хватает данных, чтобы ответить уверенно."
                    : validated.MissingData;
                sink?.OnTokenDelta(answer);
                sink?.OnAssistantMessageComplete(answer);
                return new AssistantEngineResult
                {
                    Envelope = new ResponseEnvelope
                    {
                        Kind = "clarification",
                        AnswerMarkdown = answer,
                        MissingData = answer,
                        Confidence = "high"
                    },
                    TurnResult = new TurnResult
                    {
                        StopReason = StopReason.Completed,
                        FinalAssistantText = answer,
                        AppendedItems = new List<JObject>
                        {
                            new JObject(
                                new JProperty("type", "message"),
                                new JProperty("role", "user"),
                                new JProperty("content", request.UserText ?? "")),
                            new JObject(
                                new JProperty("type", "message"),
                                new JProperty("role", "assistant"),
                                new JProperty("content", answer))
                        }
                    }
                };
            }

            if (validated.ContextMode == "snippet" || validated.ContextMode == "full")
            {
                snapshot = CaptureSnapshot(request, includeFullBodies: true);
                snapshot.Capabilities = retrieved;
            }

            var context = _contextResolver.Resolve(request, snapshot, validated);
            return await _executor.ExecuteAsync(
                request,
                context,
                validated,
                sink,
                cancellationToken).ConfigureAwait(false);
        }

        private TurnSnapshot CaptureSnapshot(TurnRequest request, bool includeFullBodies)
        {
            CurrentSelectionResult selection = null;
            try
            {
                selection = _surface.GetCurrentSelection(includeFullBodies, 1);
            }
            catch (Exception ex)
            {
                TraceLog.Write("Selection snapshot failed: " + ex.Message, "AssistantEngine");
            }

            return new TurnSnapshot
            {
                Surface = request.Surface ?? "inbox_chat",
                WorkflowId = request.WorkflowId ?? "inbox_chat",
                FolderName = !string.IsNullOrWhiteSpace(selection?.Folder)
                    ? selection.Folder
                    : request.FolderName ?? "Inbox",
                UnreadCount = request.UnreadCount,
                TotalCount = request.TotalCount,
                Selection = selection,
                Capabilities = new CapabilitySearchResult[0]
            };
        }
    }
}
