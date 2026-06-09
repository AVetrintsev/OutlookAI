using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services
{
    /// <summary>
    /// Text generation through a LiteLLM OpenAI-compatible proxy.
    /// Endpoint: {LiteLlmBaseUrl}/chat/completions.
    /// Auth: Authorization: Bearer {user-provided API key}.
    /// </summary>
    public sealed class LiteLlmChatService : IDisposable
    {
        public enum ActionType
        {
            Proofread,
            Revise,
            Draft,
            Shorten,
            Lengthen,
            Formal,
            Friendly,
            Custom
        }

        private readonly LiteLlmCredentialService _credentials;
        private readonly HttpClient _http;
        private readonly bool _ownsHttp;
        private bool _disposed;

        public LiteLlmChatService(LiteLlmCredentialService credentials)
            : this(credentials, BuildDefaultHttpClient(), ownsHttp: true)
        {
        }

        public LiteLlmChatService(LiteLlmCredentialService credentials, HttpClient httpClient, bool ownsHttp = false)
        {
            _credentials = credentials ?? throw new ArgumentNullException(nameof(credentials));
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _ownsHttp = ownsHttp;
        }

        public static string ChatCompletionsEndpoint
            => Config.NormalizeBaseUrl(Config.LiteLlmBaseUrl) + "/chat/completions";

        private static HttpClient BuildDefaultHttpClient()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            return new HttpClient { Timeout = TimeSpan.FromMinutes(2) };
        }

        public async Task<TurnResult> RunTurnAsync(
            ConversationContext context,
            string userMessage,
            IToolHost toolHost,
            ChatEventSink sink,
            CancellationToken cancellationToken)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (toolHost == null) throw new ArgumentNullException(nameof(toolHost));
            if (sink == null) sink = new ChatEventSink();

            context.History.Add(new JObject(
                new JProperty("type", "message"),
                new JProperty("role", "user"),
                new JProperty("content", userMessage ?? "")));

            var appended = new List<JObject>();
            var result = new TurnResult();

            for (int rounds = 1; rounds <= MaxToolRounds; rounds++)
            {
                result.RoundsUsed = rounds;
                var body = BuildRunTurnRequest(context);
                var assistantText = new StringBuilder();
                List<JObject> pendingCalls;
                bool roundCancelled = false;

                try
                {
                    pendingCalls = await SendChatCompletionAsync(
                        body,
                        assistantText,
                        sink,
                        cancellationToken).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    pendingCalls = new List<JObject>();
                    roundCancelled = true;
                }

                if (assistantText.Length > 0)
                {
                    var assistantItem = new JObject(
                        new JProperty("type", "message"),
                        new JProperty("role", "assistant"),
                        new JProperty("content", assistantText.ToString()));
                    context.History.Add(assistantItem);
                    appended.Add(assistantItem);
                    sink.OnAssistantMessageComplete(assistantText.ToString());
                    result.FinalAssistantText = assistantText.ToString();
                }

                if (roundCancelled)
                {
                    result.StopReason = StopReason.Cancelled;
                    result.AppendedItems = appended;
                    return result;
                }

                if (pendingCalls.Count == 0)
                {
                    sink.OnRoundBoundary();
                    result.StopReason = StopReason.Completed;
                    result.AppendedItems = appended;
                    return result;
                }

                var dispatchTasks = pendingCalls.Select(call =>
                    DispatchOneAsync(toolHost, sink, call, cancellationToken)).ToArray();

                DispatchedCall[] dispatched;
                try
                {
                    dispatched = await Task.WhenAll(dispatchTasks).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    result.StopReason = StopReason.Cancelled;
                    result.AppendedItems = appended;
                    return result;
                }

                foreach (var d in dispatched)
                {
                    context.History.Add(d.FunctionCall);
                    appended.Add(d.FunctionCall);
                }
                foreach (var d in dispatched)
                {
                    context.History.Add(d.FunctionCallOutput);
                    appended.Add(d.FunctionCallOutput);
                }

                sink.OnRoundBoundary();
            }

            result.StopReason = StopReason.MaxRoundsReached;
            result.AppendedItems = appended;
            return result;
        }

        private async Task<List<JObject>> SendChatCompletionAsync(
            JObject body,
            StringBuilder assistantText,
            ChatEventSink sink,
            CancellationToken cancellationToken)
        {
            using (var request = new HttpRequestMessage(HttpMethod.Post, ChatCompletionsEndpoint))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _credentials.GetApiKey());
                request.Headers.Accept.ParseAdd("text/event-stream");
                request.Content = new StringContent(
                    body.ToString(Newtonsoft.Json.Formatting.None),
                    Encoding.UTF8,
                    "application/json");

                using (var response = await _http.SendAsync(
                    request,
                    HttpCompletionOption.ResponseHeadersRead,
                    cancellationToken).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        var errorBody = await SafeReadAsStringAsync(response).ConfigureAwait(false);
                        sink.OnError(errorBody);
                        throw new InvalidOperationException(
                            "LiteLLM backend error: " + (int)response.StatusCode + " " + errorBody);
                    }

                    using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                    {
                        return await ReadChatCompletionsSseAsync(stream, assistantText, sink, cancellationToken)
                            .ConfigureAwait(false);
                    }
                }
            }
        }

        private static async Task<List<JObject>> ReadChatCompletionsSseAsync(
            Stream stream,
            StringBuilder assistantText,
            ChatEventSink sink,
            CancellationToken cancellationToken)
        {
            var callsByIndex = new Dictionary<int, JObject>();

            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                string line;
                while ((line = await reader.ReadLineAsync().ConfigureAwait(false)) != null)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!line.StartsWith("data:", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var payload = line.Substring(5).TrimStart();
                    if (payload == "[DONE]")
                    {
                        break;
                    }

                    JObject evt;
                    try { evt = JObject.Parse(payload); } catch { continue; }

                    var choice = evt["choices"] != null ? evt["choices"].FirstOrDefault() as JObject : null;
                    var delta = choice != null ? choice["delta"] as JObject : null;
                    if (delta == null)
                    {
                        var errorMessage = (string)evt["error"]?["message"];
                        if (!string.IsNullOrEmpty(errorMessage))
                        {
                            throw new InvalidOperationException("LiteLLM backend error: " + errorMessage);
                        }
                        continue;
                    }

                    var content = (string)delta["content"];
                    if (!string.IsNullOrEmpty(content))
                    {
                        assistantText.Append(content);
                        sink.OnTokenDelta(content);
                    }

                    var toolCalls = delta["tool_calls"] as JArray;
                    if (toolCalls == null)
                    {
                        continue;
                    }

                    foreach (var tcToken in toolCalls)
                    {
                        var tc = tcToken as JObject;
                        if (tc == null) continue;
                        var index = (int?)tc["index"] ?? 0;
                        if (!callsByIndex.TryGetValue(index, out var call))
                        {
                            call = new JObject(
                                new JProperty("type", "function_call"),
                                new JProperty("call_id", ""),
                                new JProperty("name", ""),
                                new JProperty("arguments", ""));
                            callsByIndex[index] = call;
                        }

                        var id = (string)tc["id"];
                        if (!string.IsNullOrEmpty(id))
                        {
                            call["call_id"] = id;
                        }

                        var fn = tc["function"] as JObject;
                        if (fn == null) continue;

                        var name = (string)fn["name"];
                        if (!string.IsNullOrEmpty(name))
                        {
                            call["name"] = name;
                        }

                        var argsDelta = (string)fn["arguments"];
                        if (!string.IsNullOrEmpty(argsDelta))
                        {
                            call["arguments"] = ((string)call["arguments"] ?? "") + argsDelta;
                        }
                    }
                }
            }

            var calls = callsByIndex
                .OrderBy(kv => kv.Key)
                .Select(kv => kv.Value)
                .Where(call => !string.IsNullOrEmpty((string)call["name"]))
                .ToList();

            foreach (var call in calls)
            {
                var callId = (string)call["call_id"];
                if (string.IsNullOrEmpty(callId))
                {
                    callId = "call_" + Guid.NewGuid().ToString("N");
                    call["call_id"] = callId;
                }
                sink.OnToolCallStart(callId, (string)call["name"] ?? "", (string)call["arguments"] ?? "");
            }

            return calls;
        }

        private static async Task<DispatchedCall> DispatchOneAsync(
            IToolHost toolHost,
            ChatEventSink sink,
            JObject call,
            CancellationToken ct)
        {
            var name = (string)call["name"] ?? "";
            var args = (string)call["arguments"] ?? "{}";
            var callId = (string)call["call_id"] ?? "";
            try
            {
                OutlookAI.Diagnostics.TraceLog.Write(
                    "Dispatch " + name + " call_id=" + callId
                    + " args=" + FormatTraceArgs(args),
                    "LiteLlmChat");
            }
            catch { }

            string outputJson;
            bool ok = true;
            try
            {
                outputJson = await toolHost.DispatchAsync(name, args, ct).ConfigureAwait(false);
                if (string.IsNullOrEmpty(outputJson))
                {
                    outputJson = "{}";
                }
                else if (LooksLikeErrorEnvelope(outputJson))
                {
                    ok = false;
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                outputJson = BuildErrorEnvelope(ex);
                ok = false;
            }

            sink.OnToolCallResult(callId, ok, Summarize(outputJson), outputJson);
            return new DispatchedCall
            {
                FunctionCall = new JObject(
                    new JProperty("type", "function_call"),
                    new JProperty("call_id", callId),
                    new JProperty("name", name),
                    new JProperty("arguments", args)),
                FunctionCallOutput = new JObject(
                    new JProperty("type", "function_call_output"),
                    new JProperty("call_id", callId),
                    new JProperty("output", outputJson)),
            };
        }

        private struct DispatchedCall
        {
            public JObject FunctionCall;
            public JObject FunctionCallOutput;
        }

        internal static string FormatTraceArgs(string args)
        {
            if (string.IsNullOrEmpty(args)) return args ?? "";
            const int maxTraceArgs = 500;
            return args.Length > maxTraceArgs
                ? args.Substring(0, maxTraceArgs) + "..."
                : args;
        }

        private JObject BuildRunTurnRequest(ConversationContext context)
        {
            var messages = BuildMessages(context.SystemInstructions, context.History);
            var body = BuildBaseChatBody(messages, stream: true);
            body["tools"] = ToolCatalogSchema.BuildChatCompletionsToolsArray(context.IncludeWriteTools);
            body["tool_choice"] = "auto";
            body["parallel_tool_calls"] = true;
            AddReasoningEffort(body, context.ReasoningEffortOverride);
            return body;
        }

        private static JArray BuildMessages(string systemInstructions, IEnumerable<JObject> history)
        {
            var messages = new JArray();
            if (!string.IsNullOrWhiteSpace(systemInstructions))
            {
                messages.Add(new JObject(
                    new JProperty("role", "system"),
                    new JProperty("content", systemInstructions)));
            }

            var items = (history ?? Enumerable.Empty<JObject>()).ToList();
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var type = (string)item["type"];
                if (type == "message")
                {
                    messages.Add(new JObject(
                        new JProperty("role", (string)item["role"] ?? "user"),
                        new JProperty("content", ExtractText(item["content"]))));
                }
                else if (type == "function_call")
                {
                    var toolCalls = new JArray();
                    while (i < items.Count && (string)items[i]["type"] == "function_call")
                    {
                        toolCalls.Add(new JObject(
                            new JProperty("id", (string)items[i]["call_id"] ?? ""),
                            new JProperty("type", "function"),
                            new JProperty("function", new JObject(
                                new JProperty("name", (string)items[i]["name"] ?? ""),
                                new JProperty("arguments", (string)items[i]["arguments"] ?? "{}")))));
                        i++;
                    }
                    i--;
                    messages.Add(new JObject(
                        new JProperty("role", "assistant"),
                        new JProperty("content", JValue.CreateNull()),
                        new JProperty("tool_calls", toolCalls)));
                }
                else if (type == "function_call_output")
                {
                    messages.Add(new JObject(
                        new JProperty("role", "tool"),
                        new JProperty("tool_call_id", (string)item["call_id"] ?? ""),
                        new JProperty("content", (string)item["output"] ?? "{}")));
                }
            }

            return messages;
        }

        private static JObject BuildBaseChatBody(JArray messages, bool stream)
        {
            return new JObject(
                new JProperty("model", Config.Model),
                new JProperty("messages", messages),
                new JProperty("temperature", Config.Temperature),
                new JProperty("max_tokens", Config.MaxTokens),
                new JProperty("stream", stream));
        }

        private static void AddReasoningEffort(JObject body, string overrideEffort)
        {
            var effort = !string.IsNullOrEmpty(overrideEffort)
                ? overrideEffort
                : Config.ReasoningEffort;
            if (!string.IsNullOrEmpty(effort)
                && !string.Equals(effort, "None", StringComparison.OrdinalIgnoreCase))
            {
                body["reasoning_effort"] = effort.ToLowerInvariant();
            }
        }

        public async Task<string> ProcessEmailAsync(
            ActionType action,
            string emailContent,
            string customPrompt = "",
            CancellationToken cancellationToken = default(CancellationToken))
        {
            var messages = new JArray(
                new JObject(
                    new JProperty("role", "system"),
                    new JProperty("content", GetSystemPrompt(action))),
                new JObject(
                    new JProperty("role", "user"),
                    new JProperty("content", BuildUserMessage(action, emailContent, customPrompt ?? ""))));
            var body = BuildBaseChatBody(messages, stream: true);

            var output = new StringBuilder();
            await SendChatCompletionAsync(body, output, new ChatEventSink(), cancellationToken).ConfigureAwait(false);
            return output.ToString().Trim();
        }

        private static string ExtractText(JToken token)
        {
            if (token == null) return "";
            if (token.Type == JTokenType.String) return (string)token;
            if (token.Type == JTokenType.Array)
            {
                var sb = new StringBuilder();
                foreach (var part in (JArray)token)
                {
                    var text = (string)part["text"];
                    if (!string.IsNullOrEmpty(text)) sb.Append(text);
                }
                return sb.ToString();
            }
            return token.ToString();
        }

        private static async Task<string> SafeReadAsStringAsync(HttpResponseMessage response)
        {
            try
            {
                return await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            }
            catch
            {
                return "";
            }
        }

        private static bool LooksLikeErrorEnvelope(string outputJson)
        {
            return outputJson.IndexOf("\"error\"", StringComparison.Ordinal) >= 0;
        }

        private static string BuildErrorEnvelope(Exception ex)
        {
            var err = new JObject(
                new JProperty("error", new JObject(
                    new JProperty("code", ex.GetType().Name),
                    new JProperty("message", ex.Message ?? ""))));
            return err.ToString(Newtonsoft.Json.Formatting.None);
        }

        private static string Summarize(string outputJson)
        {
            if (string.IsNullOrEmpty(outputJson)) return "";
            const int max = 120;
            var s = outputJson.Replace('\n', ' ').Replace('\r', ' ');
            return s.Length <= max ? s : s.Substring(0, max) + "...";
        }

        private const int MaxToolRounds = 16;

        public static string GetSystemPrompt(ActionType action)
        {
            switch (action)
            {
                case ActionType.Proofread:
                    return "You are a professional editor. Review the email for grammar, spelling, punctuation, and clarity issues. Return the corrected email text only. Do not add any explanations.";
                case ActionType.Revise:
                    return "You are a professional writing assistant. Improve the email clarity, flow, and impact. Return only the revised email text without any explanations.";
                case ActionType.Draft:
                    return "You are a professional email writer. Write a clear, professional email based on the instructions. If replying to an email thread, write only your reply - do not include the previous messages. Return only the email text you are composing.";
                case ActionType.Shorten:
                    return "You are a professional editor. Condense this email to be more concise while keeping essential information. Return only the shortened email text.";
                case ActionType.Lengthen:
                    return "You are a professional writer. Expand this email with more detail while maintaining professionalism. Return only the expanded email text.";
                case ActionType.Formal:
                    return "You are a professional editor. Rewrite this email in a more formal tone suitable for business. Return only the rewritten email text.";
                case ActionType.Friendly:
                    return "You are a professional editor. Rewrite this email in a warmer, friendlier tone while remaining professional. Return only the rewritten email text.";
                case ActionType.Custom:
                default:
                    return "You are a professional email writing assistant. Help the user with their email based on their instructions. Return only the result.";
            }
        }

        public static string BuildUserMessage(ActionType action, string emailContent, string customPrompt)
        {
            if (action == ActionType.Draft)
            {
                if (!string.IsNullOrWhiteSpace(emailContent))
                {
                    return "Write a reply email based on these instructions:\n\n" + customPrompt +
                           "\n\n--- Email thread for context (do NOT include this in your response, just use it for context) ---\n\n" + emailContent;
                }
                return "Write an email based on these instructions:\n\n" + customPrompt;
            }
            if (action == ActionType.Custom)
            {
                return "Email content:\n\n" + emailContent + "\n\nInstructions: " + customPrompt;
            }
            return "Email to " + action.ToString().ToLowerInvariant() + ":\n\n" + emailContent;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (_ownsHttp)
            {
                _http.Dispose();
            }
        }
    }
}
