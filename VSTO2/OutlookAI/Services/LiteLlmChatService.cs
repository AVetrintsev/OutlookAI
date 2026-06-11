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
        {
            get
            {
                Config.EnsureLiteLlmServerConfigured();
                return Config.NormalizeBaseUrl(Config.LiteLlmBaseUrl) + "/chat/completions";
            }
        }

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
            var allowedToolNames = context.AllowedToolNames == null
                ? null
                : new HashSet<string>(
                    context.AllowedToolNames.Where(name => !string.IsNullOrWhiteSpace(name)),
                    StringComparer.OrdinalIgnoreCase);

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
                    DispatchOneAsync(toolHost, sink, call, allowedToolNames, cancellationToken)).ToArray();

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
            var requestId = LlmDebugLogger.NewRequestId();
            var endpoint = ChatCompletionsEndpoint;
            LlmDebugLogger.Write(requestId, "request.metadata",
                "endpoint=" + endpoint
                + "\r\nmodel=" + Config.Model
                + "\r\ntemperature=" + Config.Temperature.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + "\r\nmax_tokens=" + Config.MaxTokens
                + "\r\nreasoning_effort=" + (Config.ReasoningEffort ?? "")
                + "\r\nstream=true"
                + "\r\napi_key_logged=false");
            LlmDebugLogger.WriteJson(requestId, "request.body", body);

            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
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
                        LlmDebugLogger.Write(requestId, "response.status",
                            "status_code=" + (int)response.StatusCode
                            + "\r\nreason=" + response.ReasonPhrase
                            + "\r\ncontent_type=" + (response.Content?.Headers?.ContentType?.ToString() ?? ""));

                        if (!response.IsSuccessStatusCode)
                        {
                            var errorBody = await SafeReadAsStringAsync(response).ConfigureAwait(false);
                            LlmDebugLogger.Write(requestId, "response.error_body", errorBody);
                            sink.OnError(errorBody);
                            throw new InvalidOperationException(
                                "LiteLLM backend error: " + (int)response.StatusCode + " " + errorBody);
                        }

                        using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                        {
                            return await ReadChatCompletionsSseAsync(stream, assistantText, sink, cancellationToken, requestId)
                                .ConfigureAwait(false);
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                LlmDebugLogger.Write(requestId, "request.cancelled", "Operation cancelled.");
                throw;
            }
            catch (Exception ex)
            {
                LlmDebugLogger.Write(requestId, "request.exception", ex.ToString());
                throw;
            }
        }

        private static async Task<List<JObject>> ReadChatCompletionsSseAsync(
            Stream stream,
            StringBuilder assistantText,
            ChatEventSink sink,
            CancellationToken cancellationToken,
            string requestId = null)
        {
            var callsByIndex = new Dictionary<int, JObject>();
            var bufferedText = new StringBuilder();
            var bufferingPotentialTextToolCall = true;
            var rawSse = new StringBuilder();

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
                    rawSse.AppendLine(payload);
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
                        if (bufferingPotentialTextToolCall)
                        {
                            bufferedText.Append(content);
                            if (!CouldStillBeTextToolCall(bufferedText.ToString()))
                            {
                                sink.OnTokenDelta(bufferedText.ToString());
                                bufferedText.Clear();
                                bufferingPotentialTextToolCall = false;
                            }
                        }
                        else
                        {
                            sink.OnTokenDelta(content);
                        }
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

            LlmDebugLogger.Write(requestId, "response.raw_sse_data", rawSse.ToString());

            if (calls.Count == 0 && TryParseTextToolCalls(assistantText.ToString(), out var textCalls))
            {
                assistantText.Clear();
                bufferedText.Clear();
                calls = textCalls;
            }
            else if (calls.Count == 0 && TryUnwrapStructuredAssistantText(assistantText.ToString(), out var unwrappedText))
            {
                assistantText.Clear();
                assistantText.Append(unwrappedText);
                bufferedText.Clear();
                sink.OnTokenDelta(unwrappedText);
            }
            else if (calls.Count == 0 && TryBuildFallbackFromRawToolJson(assistantText.ToString(), out var fallbackText))
            {
                assistantText.Clear();
                assistantText.Append(fallbackText);
                bufferedText.Clear();
                sink.OnTokenDelta(fallbackText);
            }
            else if (bufferedText.Length > 0)
            {
                sink.OnTokenDelta(bufferedText.ToString());
                bufferedText.Clear();
            }

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

            LlmDebugLogger.Write(requestId, "response.parsed",
                "assistant_text:\r\n" + assistantText
                + "\r\n\r\ntool_calls:\r\n"
                + new JArray(calls.Select(c => c.DeepClone())).ToString(Newtonsoft.Json.Formatting.Indented));

            return calls;
        }

        private static bool CouldStillBeTextToolCall(string text)
        {
            var trimmed = (text ?? "").TrimStart();
            if (trimmed.Length == 0)
            {
                return true;
            }

            if (trimmed.StartsWith("{", StringComparison.Ordinal)
                || trimmed.StartsWith("[", StringComparison.Ordinal)
                || trimmed.StartsWith("```", StringComparison.Ordinal))
            {
                return true;
            }

            if (trimmed.Equals("PDF", StringComparison.OrdinalIgnoreCase)
                || trimmed.Equals("JSON", StringComparison.OrdinalIgnoreCase)
                || trimmed.Equals("Tool", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (trimmed.StartsWith("PDF", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("JSON", StringComparison.OrdinalIgnoreCase)
                || trimmed.StartsWith("Tool", StringComparison.OrdinalIgnoreCase))
            {
                var afterFirstLine = trimmed.Split(new[] { '\r', '\n' }, 2, StringSplitOptions.RemoveEmptyEntries);
                return afterFirstLine.Length == 1
                    || afterFirstLine[1].TrimStart().StartsWith("{", StringComparison.Ordinal)
                    || afterFirstLine[1].TrimStart().StartsWith("```", StringComparison.Ordinal);
            }

            return false;
        }

        internal static bool TryParseTextToolCalls(string text, out List<JObject> calls)
        {
            calls = new List<JObject>();
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            foreach (var candidate in EnumerateJsonCandidates(text))
            {
                try
                {
                    var token = JToken.Parse(candidate);
                    if (TryConvertTextToolToken(token, out var parsedCalls))
                    {
                        calls = parsedCalls;
                        return calls.Count > 0;
                    }
                }
                catch
                {
                    // Keep scanning; small local models often wrap JSON in labels.
                }
            }

            return false;
        }

        internal static bool TryUnwrapStructuredAssistantText(string text, out string unwrappedText)
        {
            unwrappedText = "";
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            foreach (var candidate in EnumerateJsonCandidates(text))
            {
                try
                {
                    var obj = JToken.Parse(candidate) as JObject;
                    if (obj == null || obj["messages"] != null || obj["body_plaintext"] != null)
                    {
                        continue;
                    }

                    var knownText = (string)obj["answer"]
                        ?? (string)obj["response"]
                        ?? (string)obj["text"]
                        ?? (string)obj["summary"]
                        ?? (string)obj["content"];
                    if (!string.IsNullOrWhiteSpace(knownText))
                    {
                        unwrappedText = knownText.Trim();
                        return true;
                    }

                    var stringProperties = obj.Properties()
                        .Where(p => p.Value.Type == JTokenType.String)
                        .ToList();
                    if (stringProperties.Count == 1)
                    {
                        var key = stringProperties[0].Name.Trim();
                        var value = ((string)stringProperties[0].Value ?? "").Trim();
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            unwrappedText = value;
                            return true;
                        }
                        if (!string.IsNullOrWhiteSpace(key))
                        {
                            unwrappedText = key;
                            return true;
                        }
                    }

                    if (stringProperties.Count > 1)
                    {
                        var values = stringProperties
                            .Select(p => ((string)p.Value ?? "").Trim())
                            .Where(v => !string.IsNullOrWhiteSpace(v))
                            .Distinct(StringComparer.Ordinal)
                            .ToList();
                        if (values.Count == 1)
                        {
                            unwrappedText = values[0];
                            return true;
                        }
                    }
                }
                catch
                {
                    // Keep scanning; local models often prepend labels like "PDF".
                }
            }

            return false;
        }

        internal static bool TryBuildFallbackFromRawToolJson(string text, out string fallbackText)
        {
            fallbackText = "";
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            foreach (var candidate in EnumerateJsonCandidates(text))
            {
                try
                {
                    var token = JToken.Parse(candidate);
                    if (TrySummarizeOutlookToolResult(token, out fallbackText))
                    {
                        return true;
                    }
                }
                catch
                {
                    // Keep scanning; local models often prepend labels like "PDF".
                }
            }

            return false;
        }

        private static bool TrySummarizeOutlookToolResult(JToken token, out string summary)
        {
            summary = "";
            var obj = token as JObject;
            if (obj == null)
            {
                return false;
            }

            var messages = obj["messages"] as JArray;
            if (messages != null)
            {
                summary = SummarizeMessagesResult(obj, messages);
                return !string.IsNullOrWhiteSpace(summary);
            }

            if (obj["body_plaintext"] != null || obj["subject"] != null)
            {
                summary = SummarizeMessagesResult(
                    new JObject(new JProperty("count", 1)),
                    new JArray(obj));
                return !string.IsNullOrWhiteSpace(summary);
            }

            return false;
        }

        private static string SummarizeMessagesResult(JObject root, JArray messages)
        {
            var count = (int?)root["count"] ?? messages.Count;
            var sb = new StringBuilder();
            sb.AppendLine("Сводка выбранной переписки");
            sb.AppendLine();
            sb.AppendLine("Найдено сообщений: " + count + ".");

            var first = messages.OfType<JObject>().FirstOrDefault();
            if (first == null)
            {
                return sb.ToString().Trim();
            }

            var subject = CleanInline((string)first["subject"]);
            var from = CleanInline((string)first["from"]);
            var received = CleanInline((string)first["received_at"]);
            var body = CleanBody((string)first["body_plaintext"] ?? (string)first["snippet"]);

            if (!string.IsNullOrWhiteSpace(subject))
            {
                sb.AppendLine("Тема: " + subject);
            }
            if (!string.IsNullOrWhiteSpace(from))
            {
                sb.AppendLine("Отправитель: " + from);
            }
            if (!string.IsNullOrWhiteSpace(received))
            {
                sb.AppendLine("Дата: " + received);
            }

            if (!string.IsNullOrWhiteSpace(body))
            {
                sb.AppendLine();
                sb.AppendLine("Кратко:");
                foreach (var sentence in TakeSummarySentences(body, maxSentences: 4))
                {
                    sb.AppendLine("- " + sentence);
                }
            }

            if (messages.Count > 1)
            {
                sb.AppendLine();
                sb.AppendLine("В переписке несколько сообщений; выше приведена сводка первого выбранного сообщения.");
            }

            return sb.ToString().Trim();
        }

        private static IEnumerable<string> TakeSummarySentences(string text, int maxSentences)
        {
            var normalized = CleanBody(text);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                yield break;
            }

            var sentences = normalized
                .Split(new[] { ". ", "! ", "? ", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim().Trim('.', '!', '?'))
                .Where(x => x.Length > 0)
                .Take(maxSentences);

            foreach (var sentence in sentences)
            {
                yield return sentence.Length <= 260 ? sentence : sentence.Substring(0, 260).TrimEnd() + "...";
            }
        }

        private static string CleanBody(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            var normalized = text
                .Replace("\r", "\n")
                .Replace("\t", " ");
            while (normalized.Contains("\n\n\n"))
            {
                normalized = normalized.Replace("\n\n\n", "\n\n");
            }
            while (normalized.Contains("  "))
            {
                normalized = normalized.Replace("  ", " ");
            }
            return normalized.Trim();
        }

        private static string CleanInline(string text)
        {
            return string.IsNullOrWhiteSpace(text)
                ? ""
                : text.Replace("\r", " ").Replace("\n", " ").Trim();
        }

        private static IEnumerable<string> EnumerateJsonCandidates(string text)
        {
            var trimmed = (text ?? "").Trim();
            if (trimmed.StartsWith("```", StringComparison.Ordinal))
            {
                var firstNewline = trimmed.IndexOf('\n');
                var lastFence = trimmed.LastIndexOf("```", StringComparison.Ordinal);
                if (firstNewline >= 0 && lastFence > firstNewline)
                {
                    yield return trimmed.Substring(firstNewline + 1, lastFence - firstNewline - 1).Trim();
                }
            }

            yield return trimmed;

            for (int start = 0; start < text.Length; start++)
            {
                if (text[start] != '{' && text[start] != '[')
                {
                    continue;
                }

                var candidate = TryReadBalancedJson(text, start);
                if (!string.IsNullOrEmpty(candidate))
                {
                    yield return candidate;
                }
            }
        }

        private static string TryReadBalancedJson(string text, int start)
        {
            var depth = 0;
            var inString = false;
            var escaped = false;
            for (int i = start; i < text.Length; i++)
            {
                var ch = text[i];
                if (inString)
                {
                    if (escaped)
                    {
                        escaped = false;
                    }
                    else if (ch == '\\')
                    {
                        escaped = true;
                    }
                    else if (ch == '"')
                    {
                        inString = false;
                    }
                    continue;
                }

                if (ch == '"')
                {
                    inString = true;
                }
                else if (ch == '{' || ch == '[')
                {
                    depth++;
                }
                else if (ch == '}' || ch == ']')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return text.Substring(start, i - start + 1);
                    }
                }
            }

            return null;
        }

        private static bool TryConvertTextToolToken(JToken token, out List<JObject> calls)
        {
            calls = new List<JObject>();
            var arr = token as JArray;
            if (arr != null)
            {
                foreach (var item in arr)
                {
                    if (!TryConvertTextToolObject(item as JObject, out var call))
                    {
                        return false;
                    }
                    calls.Add(call);
                }
                return calls.Count > 0;
            }

            if (TryConvertTextToolObject(token as JObject, out var singleCall))
            {
                calls.Add(singleCall);
                return true;
            }

            return false;
        }

        private static bool TryConvertTextToolObject(JObject obj, out JObject call)
        {
            call = null;
            if (obj == null)
            {
                return false;
            }

            var functionObject = obj["function"] as JObject;
            var name = (string)obj["function"]
                ?? (string)obj["name"]
                ?? (string)obj["tool"]
                ?? (string)functionObject?["name"];

            if (string.IsNullOrWhiteSpace(name) || !name.StartsWith("outlook_", StringComparison.Ordinal))
            {
                return false;
            }

            var argsToken = obj["parameters"]
                ?? obj["arguments"]
                ?? obj["args"]
                ?? functionObject?["arguments"]
                ?? new JObject();

            string args;
            if (argsToken.Type == JTokenType.String)
            {
                args = ((string)argsToken) ?? "{}";
                if (string.IsNullOrWhiteSpace(args))
                {
                    args = "{}";
                }
            }
            else
            {
                args = argsToken.ToString(Newtonsoft.Json.Formatting.None);
            }

            call = new JObject(
                new JProperty("type", "function_call"),
                new JProperty("call_id", "call_" + Guid.NewGuid().ToString("N")),
                new JProperty("name", name.Trim()),
                new JProperty("arguments", args));
            return true;
        }

        private static async Task<DispatchedCall> DispatchOneAsync(
            IToolHost toolHost,
            ChatEventSink sink,
            JObject call,
            ISet<string> allowedToolNames,
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
            LlmDebugLogger.Write(callId, "tool.dispatch.start",
                "name=" + name
                + "\r\ncall_id=" + callId
                + "\r\narguments:\r\n" + (args ?? ""));

            string outputJson;
            bool ok = true;
            if (allowedToolNames != null && !allowedToolNames.Contains(name))
            {
                outputJson = new JObject(
                    new JProperty("error", new JObject(
                        new JProperty("code", "tool_not_allowed"),
                        new JProperty("message", "Tool is not allowed for this custom action."))))
                    .ToString(Newtonsoft.Json.Formatting.None);
                ok = false;
            }
            else
            {
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
                    LlmDebugLogger.Write(callId, "tool.dispatch.exception", ex.ToString());
                }
            }

            LlmDebugLogger.Write(callId, "tool.dispatch.result",
                "name=" + name
                + "\r\ncall_id=" + callId
                + "\r\nok=" + ok
                + "\r\noutput:\r\n" + outputJson);
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
            var tools = ToolCatalogSchema.BuildChatCompletionsToolsArray(
                context.IncludeWriteTools,
                context.AllowedToolNames);
            if (tools.Count > 0)
            {
                body["tools"] = tools;
                body["tool_choice"] = "auto";
            }
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
                    messages.Add(new JObject(
                        new JProperty("role", "user"),
                        new JProperty("content",
                            "Используй результат инструмента выше, чтобы ответить на исходный запрос пользователя. "
                            + "Не повторяй JSON и не показывай служебные поля. Если пользователь просил PDF или отчет, сначала подготовь понятный текст/markdown, затем при необходимости вызови подходящий export-инструмент.")));
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

        public async Task<string> CompleteWithoutToolsAsync(
            string systemInstructions,
            string userMessage,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            var messages = new JArray(
                new JObject(
                    new JProperty("role", "system"),
                    new JProperty("content", systemInstructions ?? "")),
                new JObject(
                    new JProperty("role", "user"),
                    new JProperty("content", userMessage ?? "")));
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
            return PromptCatalog.Default.Get(ActionPromptId(action));
        }

        private static string ActionPromptId(ActionType action)
        {
            return action.ToString().ToLowerInvariant();
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
