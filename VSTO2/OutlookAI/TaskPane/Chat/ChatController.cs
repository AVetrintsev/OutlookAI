using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.CustomActions;
using OutlookAI.Services.Export;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;

namespace OutlookAI.TaskPane.Chat
{
    /// <summary>
    /// Owns the Chat tab's WebView2 lifecycle and the JS&#x2194;C# bridge that
    /// wires user input into <see cref="LiteLlmChatService.RunTurnAsync"/>.
    /// Per-Inspector instance, constructed by <see cref="AITaskPane"/> after
    /// <see cref="AITaskPane.Bind"/> hands it the tool host + surface.
    /// </summary>
    public sealed class ChatController : IDisposable
    {
        private readonly Control _hostContainer;
        private readonly LiteLlmChatService _chat;
        private readonly IToolHost _toolHost;
        private readonly LiveOutlookSurface _surface;
        private readonly ConversationStore _store;
        private readonly Func<string> _composerSystemPrompt;
        private readonly ExportBridge _exportBridge;
        private readonly CustomActionStore _customActionStore = new CustomActionStore();

        private WebView2 _webView;
        private CancellationTokenSource _activeCts;
        private bool _isReady;
        private bool _isDisposed;
        private bool _turnInFlight;
        private int _nextMessageId;
        private Label _fallbackLabel;

        public ChatController(
            Control hostContainer,
            LiteLlmChatService chat,
            IToolHost toolHost,
            LiveOutlookSurface surface,
            ConversationStore store)
        {
            _hostContainer = hostContainer ?? throw new ArgumentNullException(nameof(hostContainer));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
            _surface = surface;
            _store = store ?? new ConversationStore();
            if (_surface != null)
            {
                _exportBridge = new ExportBridge(_surface, CreateExportPathPolicy(), RunScript);
            }
            _composerSystemPrompt = () => PromptCatalog.Default.Get("compose_chat");
        }

        /// <summary>
        /// Initialize the WebView2, extract WebUI resources, wire bridge.
        /// Awaited by <see cref="AITaskPane.Bind"/>. On failure, leaves a
        /// friendly fallback label on the tab.
        /// </summary>
        public async Task InitializeAsync()
        {
            TraceLog.Write(">> InitializeAsync (sync prefix)", "ChatController");
            if (!WebView2Bootstrap.IsRuntimeInstalled())
            {
                TraceLog.Write("WebView2 runtime NOT installed; showing fallback", "ChatController");
                ShowFallback("Среда WebView2 Runtime не установлена.\r\nЗапустите установщик или скачайте её:\r\n" +
                             "https://developer.microsoft.com/microsoft-edge/webview2/");
                return;
            }
            TraceLog.Write("WebView2 runtime detected; constructing WebView2 control", "ChatController");

            var themeBackground = OfficeThemeDetector.GetBackgroundColor();
            _hostContainer.BackColor = themeBackground;
            _webView = new WebView2
            {
                Dock = DockStyle.Fill,
                DefaultBackgroundColor = themeBackground
            };
            TraceLog.Write("WebView2 control constructed", "ChatController");
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_webView);
            TraceLog.Write("WebView2 control added to host container; about to await Bootstrap.InitializeAsync", "ChatController");

            try
            {
                await WebView2Bootstrap.InitializeAsync(_webView);
                TraceLog.Write("Bootstrap.InitializeAsync awaited OK", "ChatController");
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                TraceLog.Write("WebMessageReceived subscribed; navigating", "ChatController");
                _webView.CoreWebView2.Navigate("https://" + WebView2Bootstrap.VirtualHost + "/index.html");
                TraceLog.Write("Navigate called", "ChatController");
            }
            catch (Exception ex)
            {
                TraceLog.Write("InitializeAsync EXCEPTION: " + ex, "ChatController");
                System.Diagnostics.Debug.WriteLine("ChatController.InitializeAsync: " + ex);
                ShowFallback("Не удалось инициализировать WebView2: " + ex.Message);
            }
        }

        private void ShowFallback(string message)
        {
            if (_fallbackLabel != null)
            {
                _fallbackLabel.Text = message;
                return;
            }
            _fallbackLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkSlateGray,
                Text = message
            };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_fallbackLabel);
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.TryGetWebMessageAsString();
                TraceLog.Write("WebMessageReceived: " + (json?.Length > 80 ? json.Substring(0, 80) + "..." : json), "ChatController");
                if (string.IsNullOrEmpty(json)) return;
                var obj = JObject.Parse(json);
                var type = (string)obj["type"] ?? "";
                var payload = obj["payload"] as JObject;
                _ = HandleHostMessageAsync(type, payload);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnWebMessageReceived EXCEPTION: " + ex, "ChatController");
                System.Diagnostics.Debug.WriteLine("ChatController bridge parse error: " + ex);
            }
        }

        private async Task HandleHostMessageAsync(string type, JObject payload)
        {
            try
            {
                if (_exportBridge != null && await _exportBridge.HandleAsync(
                    type, payload, _activeCts?.Token ?? CancellationToken.None).ConfigureAwait(false))
                {
                    return;
                }

                switch (type)
                {
                    case "ready":
                        OnWebViewReady();
                        break;
                    case "send":
                        _ = StartTurnAsync(
                            (string)payload?["text"] ?? "",
                            (string)payload?["reasoning"]);
                        break;
                    case "custom_action":
                        _ = StartCustomActionAsync((string)payload?["id"] ?? "");
                        break;
                    case "custom_action_create":
                        SaveCustomAction(payload);
                        PushCustomActionChips();
                        break;
                    case "custom_action_delete":
                        DeleteCustomAction((string)payload?["id"] ?? "");
                        PushCustomActionChips();
                        break;
                    case "stop":
                        try { _activeCts?.Cancel(); } catch { }
                        break;
                    case "clear":
                        _store.Clear();
                        _ = RunScript("outlookai.clear();");
                        break;
                    case "copy":
                        var clip = _store.ExportForClipboard();
                        try { Clipboard.SetText(clip ?? ""); } catch { /* clipboard occasionally throws on Outlook */ }
                        break;
                    case "theme_request":
                        PushTheme();
                        break;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("HandleHostMessageAsync EXCEPTION: " + ex, "ChatController");
            }
        }

        private static IExportPathPolicy CreateExportPathPolicy()
            => Globals.ThisAddIn?.ExportPathPolicy ?? new ExportPathPolicy(new ExportPathResolver());

        private void OnWebViewReady()
        {
            TraceLog.Write("OnWebViewReady entered", "ChatController");
            _isReady = true;
            PushTheme();
            PushReasoningOptions();
            PushContextStripFromSurface();
            PushCustomActionChips();
            TraceLog.Write("OnWebViewReady completed", "ChatController");
        }

        private void PushTheme()
        {
            _ = RunScript("outlookai.applyTheme(" +
                JsString(OfficeThemeDetector.GetThemeName()) + ");");
        }

        private void PushCustomActionChips()
        {
            try
            {
                var chipsArr = new JArray();
                foreach (var custom in _customActionStore.Load())
                {
                    chipsArr.Add(new JObject(
                        new JProperty("id", custom.Id),
                        new JProperty("type", "custom_action"),
                        new JProperty("label", custom.Title),
                        new JProperty("prompt", custom.Description ?? custom.Prompt ?? ""),
                        new JProperty("description", custom.Description ?? ""),
                        new JProperty("action_prompt", custom.Prompt ?? ""),
                        new JProperty("source", custom.Context?.Source ?? "current_selection"),
                        new JProperty("read_filter", custom.Context?.ReadFilter ?? "all"),
                        new JProperty("time_range", custom.Context?.TimeRange ?? "today"),
                        new JProperty("manual_from", custom.Context?.ManualFrom?.ToString("o")),
                        new JProperty("manual_to", custom.Context?.ManualTo?.ToString("o")),
                        new JProperty("max_items", custom.Context?.MaxItems ?? 20),
                        new JProperty("include_full_bodies", custom.Context?.IncludeFullBodies ?? true),
                        new JProperty("include_attachments", custom.Context?.IncludeAttachments ?? false),
                        new JProperty("output", custom.Output ?? "chat"),
                        new JProperty("allowed_tools", new JArray(custom.AllowedTools ?? new string[0]))));
                }
                var options = new JObject(
                    new JProperty("allowCustomActionManagement", true),
                    new JProperty("customActionTools", ToolManifestCatalog.Default.BuildUiToolsArray()));
                _ = RunScript("outlookai.setQuickActions(" +
                    chipsArr.ToString(Newtonsoft.Json.Formatting.None) +
                    ", " + options.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushCustomActionChips error: " + ex.Message, "ChatController");
            }
        }

        /// <summary>
        /// Push the model-aware reasoning-effort options to the WebUI.
        /// Computed from Config.ReasoningEffortsForModel(Config.Model) so the
        /// dropdown matches what the current model actually accepts (e.g.
        /// gpt-5.5 rejects 'minimal'; gpt-4.1-mini accepts only 'none').
        /// </summary>
        private void PushReasoningOptions()
        {
            try
            {
                var efforts = Config.ReasoningEffortsForModel(Config.Model);
                var arr = new JArray();
                foreach (var e in efforts) arr.Add(e);
                TraceLog.Write("Push reasoning options for " + Config.Model + ": " + string.Join(",", efforts), "ChatController");
                _ = RunScript("outlookai.setReasoningOptions(" +
                    arr.ToString(Newtonsoft.Json.Formatting.None) + ", '');");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushReasoningOptions error: " + ex.Message, "ChatController");
            }
        }

        private void PushContextStripFromSurface()
        {
            if (_surface == null) { TraceLog.Write("PushContextStrip: _surface is null", "ChatController"); return; }
            try
            {
                TraceLog.Write(">> PushContextStrip calling GetCurrentComposeState", "ChatController");
                var state = _surface.GetCurrentComposeState(includeFullBody: false);
                TraceLog.Write("<< PushContextStrip GetCurrentComposeState returned", "ChatController");
                var ctx = new JObject(
                    new JProperty("subject", state?.Subject ?? ""),
                    new JProperty("recipients", new JArray((state?.ToRecipients ?? new string[0]).Take(3))),
                    new JProperty("thread", state?.InReplyTo?.ThreadTopic ?? ""));
                _ = RunScript("outlookai.setContextStrip(" + ctx.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushContextStrip EXCEPTION: " + ex, "ChatController");
                System.Diagnostics.Debug.WriteLine("PushContextStrip error: " + ex);
            }
        }

        private async Task StartTurnAsync(string userText, string reasoningOverride)
        {
            TraceLog.Write(">> StartTurnAsync inFlight=" + _turnInFlight + " ready=" + _isReady, "ChatController");
            if (_turnInFlight || string.IsNullOrWhiteSpace(userText) || !_isReady)
            {
                TraceLog.Write("StartTurnAsync aborted (gate)", "ChatController");
                return;
            }
            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();

            await RunScript("outlookai.appendUserMessage(" + JsString(userText) + ");");
            await RunScript("outlookai.setComposerEnabled(false, true);");
            var assistantId = "asst_" + (++_nextMessageId);
            await RunScript("outlookai.appendAssistantMessage(" + JsString(assistantId) + ", '');");

            try
            {
                // RunTurnAsync mutates context.History in place (adds the user
                // message, then assistant/function-call/output items). We start
                // with a snapshot of the existing store, let the turn evolve it,
                // and then sync the diff back into the store at the end.
                var initialSnapshot = _store.Snapshot();
                var ctx = new ConversationContext
                {
                    SystemInstructions = BuildSystemInstructionsWithComposeContext(),
                    History = new System.Collections.Generic.List<JObject>(initialSnapshot),
                    IncludeWriteTools = Config.WriteToolsEnabled,
                    ReasoningEffortOverride = string.IsNullOrEmpty(reasoningOverride) ? null : reasoningOverride
                };

                var sink = new WebViewSink(this, assistantId);
                var result = await _chat.RunTurnAsync(ctx, userText, _toolHost, sink, _activeCts.Token);

                // Sync newly-appended items (user msg + assistant + tool round-
                // trips) back into the store so the next turn starts from the
                // updated history.
                for (int i = initialSnapshot.Count; i < ctx.History.Count; i++)
                {
                    _store.Append(ctx.History[i]);
                }

                var opts = new JObject(
                    new JProperty("stopped", result.StopReason == StopReason.Cancelled),
                    new JProperty("error", result.StopReason == StopReason.Error));
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", " +
                                opts.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (OperationCanceledException)
            {
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) +
                                ", {stopped:true});");
            }
            catch (Exception ex)
            {
                TraceLog.Write("StartTurnAsync EXCEPTION: " + ex, "ChatController");
                await RunScript("outlookai.showError(" + JsString(FormatTurnError(ex)) + ");");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("outlookai.setComposerEnabled(true, false);");
                TraceLog.Write("<< StartTurnAsync", "ChatController");
            }
        }

        private string BuildSystemInstructionsWithComposeContext()
        {
            var prompt = _composerSystemPrompt();
            try
            {
                var state = _surface?.GetCurrentComposeState(includeFullBody: false);
                if (state != null)
                {
                    var sb = new System.Text.StringBuilder(prompt);
                    sb.AppendLine();
                    sb.AppendLine();
                    sb.AppendLine("---");
                    sb.AppendLine("Current compose state (read-only context):");
                    if (!string.IsNullOrEmpty(state.Subject))
                    {
                        sb.AppendLine("Subject: " + state.Subject);
                    }
                    if (state.ToRecipients != null && state.ToRecipients.Any())
                    {
                        sb.AppendLine("To: " + string.Join(", ", state.ToRecipients));
                    }
                    if (state.CcRecipients != null && state.CcRecipients.Any())
                    {
                        sb.AppendLine("Cc: " + string.Join(", ", state.CcRecipients));
                    }
                    if (state.InReplyTo != null && !string.IsNullOrEmpty(state.InReplyTo.ThreadTopic))
                    {
                        sb.AppendLine("Reply-to thread: " + state.InReplyTo.ThreadTopic);
                    }
                    if (!string.IsNullOrEmpty(state.BodyPlaintext))
                    {
                        sb.AppendLine("Body (current draft, may be empty):");
                        sb.AppendLine(state.BodyPlaintext);
                    }
                    sb.AppendLine("---");
                    return sb.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("BuildSystemInstructions compose-state error: " + ex);
            }
            return prompt;
        }

        private async Task StartCustomActionAsync(string actionId)
        {
            TraceLog.Write(">> StartCustomActionAsync id=" + actionId, "ChatController");
            if (_turnInFlight || string.IsNullOrWhiteSpace(actionId) || !_isReady || _surface == null)
            {
                TraceLog.Write("StartCustomActionAsync aborted (gate)", "ChatController");
                return;
            }

            var action = _customActionStore.Load()
                .FirstOrDefault(a => string.Equals(a.Id, actionId, StringComparison.OrdinalIgnoreCase));
            if (action == null)
            {
                await RunScript("outlookai.showError(" + JsString("Пользовательское действие не найдено.") + ");");
                return;
            }

            _turnInFlight = true;
            _activeCts = new CancellationTokenSource();
            await RunScript("outlookai.appendUserMessage(" + JsString(action.Title) + ");");
            await RunScript("outlookai.setComposerEnabled(false, true);");
            var assistantId = "asst_" + (++_nextMessageId);
            await RunScript("outlookai.appendAssistantMessage(" + JsString(assistantId) + ", '');");

            try
            {
                var runner = new CustomActionRunner(_chat, _surface, _toolHost);
                var sink = new WebViewSink(this, assistantId);
                var result = await runner.RunAsync(action, _activeCts.Token, sink).ConfigureAwait(false);
                if (!sink.HasReceivedText && !string.IsNullOrEmpty(result.Text))
                {
                    await RunScript("outlookai.appendTextDelta(" +
                        JsString(assistantId) + ", " + JsString(result.Text) + ");");
                }
                if (!string.IsNullOrWhiteSpace(result.FilePath))
                {
                    var fileInfo = new JObject(
                        new JProperty("path", result.FilePath),
                        new JProperty("format", action.Output == "export_pdf" ? "pdf" : "file"));
                    await RunScript("outlookai.onFileSaved(" +
                        JsString(assistantId) + ", " + fileInfo.ToString(Newtonsoft.Json.Formatting.None) + ");");
                }
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", {});");
            }
            catch (OperationCanceledException)
            {
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", {stopped:true});");
            }
            catch (Exception ex)
            {
                TraceLog.Write("StartCustomActionAsync EXCEPTION: " + ex, "ChatController");
                await RunScript("outlookai.showError(" + JsString(FormatTurnError(ex)) + ");");
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", {error:true});");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("outlookai.setComposerEnabled(true, false);");
                TraceLog.Write("<< StartCustomActionAsync", "ChatController");
            }
        }

        private void SaveCustomAction(JObject payload)
        {
            try
            {
                if (payload == null)
                {
                    return;
                }

                var title = ((string)payload["title"] ?? "").Trim();
                var prompt = ((string)payload["prompt"] ?? "").Trim();
                if (string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(prompt))
                {
                    _ = RunScript("outlookai.showError(" + JsString("Заполните название и промпт действия.") + ");");
                    return;
                }

                var source = (string)payload["source"] ?? "current_selection";
                var existingId = ((string)payload["id"] ?? "").Trim();
                var isFolderSource = source == "current_folder" || source == "all_folders";
                var timeRange = isFolderSource
                    ? (string)payload["time_range"] ?? "today"
                    : "today";
                var includeFullBodies = (bool?)payload["include_full_bodies"] ?? true;
                var availableTools = new System.Collections.Generic.HashSet<string>(
                    ToolManifestCatalog.Default.BuildUiToolsArray()
                        .OfType<JObject>()
                        .Select(tool => (string)tool["name"])
                        .Where(name => !string.IsNullOrWhiteSpace(name)),
                    StringComparer.OrdinalIgnoreCase);
                var allowedTools = ((payload["allowed_tools"] as JArray)?.Values<string>()
                    ?? Enumerable.Empty<string>())
                    .Where(availableTools.Contains)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                var action = new CustomActionDefinition
                {
                    Id = string.IsNullOrWhiteSpace(existingId)
                        ? CustomActionStore.MakeActionId(title)
                        : existingId,
                    Title = title,
                    Description = ((string)payload["description"] ?? "").Trim(),
                    Prompt = prompt,
                    Context = new CustomActionContext
                    {
                        Source = source,
                        MessageScope = source == "related_thread" ? "thread" : "selected",
                        FolderScope = source == "all_folders" ? "all_folders" : "current_folder",
                        ReadFilter = isFolderSource
                            ? (string)payload["read_filter"] ?? "all"
                            : "all",
                        TimeRange = timeRange,
                        ManualFrom = timeRange == "manual"
                            ? (DateTimeOffset?)payload["manual_from"]
                            : null,
                        ManualTo = timeRange == "manual"
                            ? (DateTimeOffset?)payload["manual_to"]
                            : null,
                        IncludeFullBodies = includeFullBodies,
                        IncludeAttachments = (!isFolderSource || includeFullBodies)
                            ? (bool?)payload["include_attachments"] ?? false
                            : false,
                        MaxItems = source == "current_open_message"
                            ? 1
                            : Math.Max(1, Math.Min(100, (int?)payload["max_items"] ?? 20))
                    },
                    Output = (string)payload["output"] ?? "chat",
                    AllowTools = allowedTools.Length > 0,
                    AllowedTools = allowedTools
                };

                _customActionStore.Upsert(action);
            }
            catch (Exception ex)
            {
                TraceLog.Write("SaveCustomAction error: " + ex, "ChatController");
                _ = RunScript("outlookai.showError(" + JsString("Не удалось сохранить действие: " + ex.Message) + ");");
            }
        }

        private void DeleteCustomAction(string id)
        {
            try
            {
                if (!_customActionStore.Delete(id))
                {
                    _ = RunScript("outlookai.showError(" + JsString("Пользовательское действие не найдено.") + ");");
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("DeleteCustomAction error: " + ex, "ChatController");
                _ = RunScript("outlookai.showError(" + JsString("Не удалось удалить действие: " + ex.Message) + ");");
            }
        }

        private static string FormatTurnError(Exception ex)
        {
            var detail = ex?.Message ?? "";
            if (ex?.InnerException != null && !string.IsNullOrWhiteSpace(ex.InnerException.Message))
            {
                detail += " " + ex.InnerException.Message;
            }
            if (string.IsNullOrWhiteSpace(detail))
            {
                detail = "неизвестная ошибка";
            }
            return "Ошибка при обращении к LiteLLM: " + detail;
        }

        private async Task RunScript(string script)
        {
            if (_isDisposed) return;
            try
            {
                // EVERY WebView2 property/method access (including just
                // reading _webView.CoreWebView2) MUST happen on the UI thread.
                // Marshal first, THEN touch the control.
                var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                if (marshaller != null
                    && Thread.CurrentThread.ManagedThreadId != marshaller.UiThreadId)
                {
                    await marshaller.RunAsync(() =>
                    {
                        if (_isDisposed) return;
                        var core = _webView?.CoreWebView2;
                        if (core == null) return;
                        _ = core.ExecuteScriptAsync(script);
                    }, CancellationToken.None).ConfigureAwait(false);
                    return;
                }
                // On UI thread: safe to touch the control directly.
                var coreOnUi = _webView?.CoreWebView2;
                if (coreOnUi == null) return;
                await coreOnUi.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunScript EXCEPTION: " + ex.Message, "ChatController");
                System.Diagnostics.Debug.WriteLine("ExecuteScriptAsync failed: " + ex);
            }
        }

        // JSON-encode a string for safe inlining inside an ExecuteScriptAsync
        // call. Newtonsoft handles quote / backslash / control characters.
        private static string JsString(string s)
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(s ?? "");
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            try { _activeCts?.Cancel(); } catch { }
            try { _webView?.Dispose(); } catch { }
        }

        /// <summary>
        /// Streams ChatEventSink callbacks back into the WebView2 surface via
        /// ExecuteScriptAsync. Fire-and-forget; failures are swallowed and
        /// logged because the chat loop should not stall on UI hiccups.
        /// </summary>
        private sealed class WebViewSink : ChatEventSink
        {
            private readonly ChatController _owner;
            private readonly string _assistantId;
            public bool HasReceivedText { get; private set; }
            public WebViewSink(ChatController owner, string assistantId)
            {
                _owner = owner;
                _assistantId = assistantId;
            }
            public override void OnTokenDelta(string delta)
            {
                if (!string.IsNullOrEmpty(delta)) HasReceivedText = true;
                TraceLog.Write("Sink.OnTokenDelta len=" + (delta?.Length ?? 0), "WebViewSink");
                _ = _owner.RunScript(
                    "outlookai.appendTextDelta(" +
                    JsString(_assistantId) + ", " + JsString(delta) + ");");
            }
            public override void OnToolCallStart(string callId, string name, string argsJson)
            {
                TraceLog.Write("Sink.OnToolCallStart " + name + " id=" + callId, "WebViewSink");
                _ = _owner.RunScript(
                    "outlookai.appendToolCallCard(" +
                    JsString(callId) + ", " + JsString(name) + ", " + JsString(argsJson) + ");");
            }
            public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson)
            {
                TraceLog.Write("Sink.OnToolCallResult ok=" + ok + " id=" + callId, "WebViewSink");
                _ = _owner.RunScript(
                    "outlookai.updateToolCallCard(" +
                    JsString(callId) + ", " + (ok ? "true" : "false") + ", " +
                    JsString(summary) + ", " + JsString(resultJson) + ");");
                if (IsWriteTool(callId) && ok)
                {
                    _ = _owner.RunScript(
                        "outlookai.appendAuditRow(" + JsString("Wrote: " + summary) + ");");
                }
            }
            public override void OnError(string message)
            {
                TraceLog.Write("Sink.OnError: " + message, "WebViewSink");
                _ = _owner.RunScript("outlookai.showError(" + JsString(message ?? "") + ");");
            }
            public override void OnAssistantMessageComplete(string text)
            {
                TraceLog.Write("Sink.OnAssistantMessageComplete len=" + (text?.Length ?? 0), "WebViewSink");
            }
            public override void OnRoundBoundary()
            {
                TraceLog.Write("Sink.OnRoundBoundary", "WebViewSink");
            }

            // The C# side doesn't currently track which callIds correspond
            // to write tools - the model decides at runtime. For now we
            // detect write-tool intent by name prefix as a best effort.
            private static bool IsWriteTool(string callId)
            {
                // CallId doesn't carry the tool name; this method is a stub
                // until we plumb the name through. Returning false keeps the
                // audit row off until the explicit plumbing is in place
                // (avoids false-positive audit rows for read tools).
                return false;
            }
        }
    }
}
