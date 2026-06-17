using System;
using System.Drawing;
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
using OutlookAI.TaskPane.Chat;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Linq;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// Drives the per-Explorer Inbox Copilot chat surface. Mirrors
    /// Phase 2's ChatController but is anchored to an Explorer instead
    /// of an Inspector. Builds a fresh system prompt + quick-action
    /// chip set on every selection change and on every turn.
    /// </summary>
    public sealed class InboxCopilotController : IDisposable
    {
        private readonly Control _hostContainer;
        private readonly LiteLlmChatService _chat;
        private readonly IToolHost _toolHost;
        private readonly LiveOutlookSurface _surface;
        private readonly ConversationStore _store;
        private readonly Outlook.Explorer _explorer;
        private readonly ExportBridge _exportBridge;
        private readonly CustomActionStore _customActionStore = new CustomActionStore();
        private readonly CustomActionRecommendationService _recommendationService;
        private readonly ConcurrentDictionary<string, string[]> _recommendationCache =
            new ConcurrentDictionary<string, string[]>(StringComparer.Ordinal);

        private WebView2 _webView;
        private CancellationTokenSource _activeCts;
        private CancellationTokenSource _recommendationCts;
        private bool _isReady;
        private bool _isDisposed;
        private bool _turnInFlight;
        private int _nextMessageId;
        private Label _fallbackLabel;

        public InboxCopilotController(
            Control hostContainer,
            LiteLlmChatService chat,
            IToolHost toolHost,
            LiveOutlookSurface surface,
            ConversationStore store,
            Outlook.Explorer explorer)
        {
            _hostContainer = hostContainer ?? throw new ArgumentNullException(nameof(hostContainer));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
            _surface = surface;
            _store = store ?? new ConversationStore();
            _explorer = explorer;
            _recommendationService = new CustomActionRecommendationService(_chat);
            if (_surface != null)
            {
                _exportBridge = new ExportBridge(_surface, CreateExportPathPolicy(), RunScript);
            }
        }

        public async Task InitializeAsync()
        {
            TraceLog.Write(">> InitializeAsync (sync prefix)", "InboxCopilot");
            if (!WebView2Bootstrap.IsRuntimeInstalled())
            {
                ShowFallback("Среда WebView2 Runtime не установлена.\r\nЗапустите установщик или скачайте её:\r\n" +
                             "https://developer.microsoft.com/microsoft-edge/webview2/");
                return;
            }
            var themeBackground = OfficeThemeDetector.GetBackgroundColor();
            _hostContainer.BackColor = themeBackground;
            _webView = new WebView2
            {
                Dock = DockStyle.Fill,
                DefaultBackgroundColor = themeBackground
            };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_webView);

            try
            {
                await WebView2Bootstrap.InitializeAsync(_webView);
                _webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                _webView.CoreWebView2.Navigate("https://" + WebView2Bootstrap.VirtualHost + "/index.html");
                if (_explorer != null)
                {
                    _explorer.SelectionChange += OnExplorerSelectionChange;
                    _explorer.FolderSwitch += OnExplorerFolderSwitch;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("InitializeAsync EXCEPTION: " + ex, "InboxCopilot");
                ShowFallback("Не удалось инициализировать WebView2: " + ex.Message);
            }
        }

        private void ShowFallback(string message)
        {
            if (_fallbackLabel != null) { _fallbackLabel.Text = message; return; }
            _fallbackLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkSlateGray,
                Text = message,
            };
            _hostContainer.Controls.Clear();
            _hostContainer.Controls.Add(_fallbackLabel);
        }

        private void OnExplorerSelectionChange()
        {
            if (_isDisposed || !_isReady) return;
            if (Config.RecommendationsEnabled)
            {
                _ = PushRecommendationStateAsync("loading");
            }
            PushContextStripAndChips();
            _ = RefreshRecommendationsAsync();
        }

        private void OnExplorerFolderSwitch()
        {
            if (_isDisposed || !_isReady) return;
            if (Config.RecommendationsEnabled)
            {
                _ = PushRecommendationStateAsync("loading");
            }
            PushContextStripAndChips();
            _ = RefreshRecommendationsAsync();
        }

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var json = e.TryGetWebMessageAsString();
                TraceLog.Write("WebMessageReceived: " + (json?.Length > 80 ? json.Substring(0, 80) + "..." : json), "InboxCopilot");
                if (string.IsNullOrEmpty(json)) return;
                var obj = JObject.Parse(json);
                var type = (string)obj["type"] ?? "";
                var payload = obj["payload"] as JObject;
                _ = HandleHostMessageAsync(type, payload);
            }
            catch (Exception ex)
            {
                TraceLog.Write("OnWebMessageReceived EXCEPTION: " + ex, "InboxCopilot");
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
                    case "open_item":
                        _surface?.OpenItem((string)payload?["id"] ?? "");
                        break;
                    case "custom_action_create":
                        SaveCustomAction(payload);
                        _recommendationCache.Clear();
                        PushContextStripAndChips();
                        _ = RefreshRecommendationsAsync();
                        break;
                    case "custom_action_delete":
                        DeleteCustomAction((string)payload?["id"] ?? "");
                        _recommendationCache.Clear();
                        PushContextStripAndChips();
                        _ = RefreshRecommendationsAsync();
                        break;
                    case "custom_action_reorder":
                        ReorderCustomActions(payload);
                        _recommendationCache.Clear();
                        PushContextStripAndChips();
                        _ = RefreshRecommendationsAsync();
                        break;
                    case "custom_action_reset_group":
                        _customActionStore.ResetGroup((string)payload?["group_id"] ?? "");
                        _recommendationCache.Clear();
                        PushContextStripAndChips();
                        _ = RefreshRecommendationsAsync();
                        break;
                    case "custom_action_import_preview":
                        PreviewCustomActionImport((string)payload?["yaml"] ?? "");
                        break;
                    case "custom_action_import_apply":
                        ApplyCustomActionImport((string)payload?["yaml"] ?? "");
                        _recommendationCache.Clear();
                        PushContextStripAndChips();
                        _ = RefreshRecommendationsAsync();
                        break;
                    case "custom_action_export":
                        _ = ExportCustomActionsAsync((string)payload?["id"]);
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
                        try { Clipboard.SetText(clip ?? ""); } catch { }
                        break;
                    case "theme_request":
                        PushTheme();
                        break;
                }
            }
            catch (Exception ex)
            {
                TraceLog.Write("HandleHostMessageAsync EXCEPTION: " + ex, "InboxCopilot");
            }
        }

        private static IExportPathPolicy CreateExportPathPolicy()
            => Globals.ThisAddIn?.ExportPathPolicy ?? new ExportPathPolicy(new ExportPathResolver());

        private void OnWebViewReady()
        {
            TraceLog.Write("OnWebViewReady entered", "InboxCopilot");
            _isReady = true;
            PushTheme();
            PushReasoningOptions();
            PushContextStripAndChips();
            _ = RefreshRecommendationsAsync();
            TraceLog.Write("OnWebViewReady completed", "InboxCopilot");
        }

        private void PushTheme()
        {
            _ = RunScript("outlookai.applyTheme(" +
                JsString(OfficeThemeDetector.GetThemeName()) + ");");
        }

        private void PushReasoningOptions()
        {
            try
            {
                var efforts = Config.ReasoningEffortsForModel(Config.Model);
                var arr = new JArray();
                foreach (var e in efforts) arr.Add(e);
                _ = RunScript("outlookai.setReasoningOptions(" +
                    arr.ToString(Newtonsoft.Json.Formatting.None) + ", '');");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushReasoningOptions error: " + ex.Message, "InboxCopilot");
            }
        }

        private void PushContextStripAndChips()
        {
            try
            {
                CurrentSelectionResult sel = null;
                string folderName = "";
                int unreadCount = 0, totalCount = 0;
                try
                {
                    sel = _surface.GetCurrentSelection(includeFullBodies: false, maxItems: 5);
                    folderName = sel?.Folder ?? "";
                    var folder = _explorer?.CurrentFolder;
                    try
                    {
                        if (folder != null)
                        {
                            unreadCount = folder.UnReadItemCount;
                            totalCount = folder.Items.Count;
                        }
                    }
                    catch { }
                }
                catch (Exception ex) { TraceLog.Write("PushContextStrip surface error: " + ex.Message, "InboxCopilot"); }

                var ctx = new JObject(
                    new JProperty("folder", folderName),
                    new JProperty("unread_count", unreadCount),
                    new JProperty("total_count", totalCount));
                if (sel != null && sel.Count > 0 && sel.Messages != null && sel.Messages.Count > 0)
                {
                    var first = sel.Messages[0];
                    ctx.Add("selection", new JObject(
                        new JProperty("count", sel.Count),
                        new JProperty("subject", first.Subject ?? ""),
                        new JProperty("from", first.From ?? "")));
                }
                _ = RunScript("outlookai.setContextStrip(" +
                    ctx.ToString(Newtonsoft.Json.Formatting.None) + ");");

                var payload = CustomActionUiSerializer.Build(
                    _customActionStore.LoadCatalog(),
                    null,
                    ToolManifestCatalog.Default.BuildUiToolsArray());
                _ = RunScript("outlookai.setActionCatalog(" +
                    payload.ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (Exception ex)
            {
                TraceLog.Write("PushContextStripAndChips error: " + ex.Message, "InboxCopilot");
            }
        }

        private async Task RefreshRecommendationsAsync()
        {
            if (!Config.RecommendationsEnabled)
            {
                try { _recommendationCts?.Cancel(); } catch { }
                await PushRecommendationStateAsync("disabled");
                return;
            }
            CurrentSelectionResult selection;
            try { selection = _surface?.GetCurrentSelection(includeFullBodies: true, maxItems: 3); }
            catch { return; }
            var key = CustomActionRecommendationService.CacheKey(selection);
            if (string.IsNullOrEmpty(key))
            {
                await PushRecommendationStateAsync("empty");
                return;
            }
            if (_recommendationCache.TryGetValue(key, out var cached))
            {
                await PushRecommendationStateAsync("ready", cached);
                return;
            }
            try { _recommendationCts?.Cancel(); } catch { }
            _recommendationCts?.Dispose();
            _recommendationCts = new CancellationTokenSource();
            var token = _recommendationCts.Token;
            await PushRecommendationStateAsync("loading");
            try
            {
                var ids = await _recommendationService.RecommendAsync(
                    selection,
                    _customActionStore.LoadCatalog().Groups,
                    token).ConfigureAwait(false);
                if (token.IsCancellationRequested) return;
                var current = _surface.GetCurrentSelection(includeFullBodies: false, maxItems: 3);
                if (!string.Equals(key, CustomActionRecommendationService.CacheKey(current), StringComparison.Ordinal))
                    return;
                _recommendationCache[key] = ids;
                await PushRecommendationStateAsync(ids.Length > 0 ? "ready" : "empty", ids);
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                TraceLog.Write("Recommendation error: " + ex.Message, "InboxCopilot");
                if (!token.IsCancellationRequested)
                {
                    await PushRecommendationStateAsync("error");
                }
            }
        }

        private Task PushRecommendationStateAsync(string state, string[] ids = null)
        {
            return RunScript("outlookai.setActionRecommendationState(" +
                JsString(state) + ", " +
                new JArray(ids ?? new string[0]).ToString(Newtonsoft.Json.Formatting.None) + ");");
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

                _customActionStore.Upsert((string)payload["group_id"], action);
            }
            catch (Exception ex)
            {
                TraceLog.Write("SaveCustomAction error: " + ex, "InboxCopilot");
                _ = RunScript("outlookai.showError(" + JsString("Не удалось сохранить действие: " + ex.Message) + ");");
            }
        }

        private void ReorderCustomActions(JObject payload)
        {
            _customActionStore.Reorder(
                (string)payload?["group_id"] ?? "",
                (payload?["action_ids"] as JArray)?.Values<string>().ToArray() ?? new string[0]);
        }

        private void PreviewCustomActionImport(string yaml)
        {
            try
            {
                var preview = CustomActionYamlCodec.Preview(yaml, _customActionStore.LoadCatalog());
                _ = RunScript("outlookai.showActionImportPreview(" +
                    JsString(yaml) + ", " +
                    new JArray(preview.Replacements).ToString(Newtonsoft.Json.Formatting.None) + ");");
            }
            catch (Exception ex)
            {
                _ = RunScript("outlookai.showError(" + JsString("Ошибка YAML: " + ex.Message) + ");");
            }
        }

        private void ApplyCustomActionImport(string yaml)
        {
            var preview = CustomActionYamlCodec.Preview(yaml, _customActionStore.LoadCatalog());
            _customActionStore.SaveCatalog(CustomActionYamlCodec.Merge(
                _customActionStore.LoadCatalog(),
                preview.Imported));
        }

        private async Task ExportCustomActionsAsync(string actionId)
        {
            var yaml = CustomActionYamlCodec.Export(_customActionStore.LoadCatalog(), actionId);
            Action showDialog = () =>
            {
                using (var dialog = new SaveFileDialog
                {
                    Filter = "YAML (*.yaml)|*.yaml|All files (*.*)|*.*",
                    FileName = string.IsNullOrWhiteSpace(actionId) ? "outlookai-actions.yaml" : actionId + ".yaml"
                })
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        File.WriteAllText(dialog.FileName, yaml, System.Text.Encoding.UTF8);
                    }
                }
            };
            var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
            if (marshaller != null && Thread.CurrentThread.ManagedThreadId != marshaller.UiThreadId)
            {
                await marshaller.RunAsync(showDialog, CancellationToken.None).ConfigureAwait(false);
            }
            else
            {
                showDialog();
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
                TraceLog.Write("DeleteCustomAction error: " + ex, "InboxCopilot");
                _ = RunScript("outlookai.showError(" + JsString("Не удалось удалить действие: " + ex.Message) + ");");
            }
        }

        private async Task StartTurnAsync(string userText, string reasoningOverride)
        {
            TraceLog.Write(">> StartTurnAsync inFlight=" + _turnInFlight + " ready=" + _isReady, "InboxCopilot");
            if (_turnInFlight || string.IsNullOrWhiteSpace(userText) || !_isReady)
            {
                TraceLog.Write("StartTurnAsync aborted (gate)", "InboxCopilot");
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
                var initialSnapshot = _store.Snapshot();
                var ctx = new ConversationContext
                {
                    SystemInstructions = BuildSystemInstructionsForCurrentState(),
                    History = new System.Collections.Generic.List<JObject>(initialSnapshot),
                    IncludeWriteTools = Config.WriteToolsEnabled,
                    ReasoningEffortOverride = string.IsNullOrEmpty(reasoningOverride) ? null : reasoningOverride,
                };

                var sink = new WebViewSink(this, assistantId);
                var result = await _chat.RunTurnAsync(ctx, userText, _toolHost, sink, _activeCts.Token);

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
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", {stopped:true});");
            }
            catch (Exception ex)
            {
                TraceLog.Write("StartTurnAsync EXCEPTION: " + ex, "InboxCopilot");
                await RunScript("outlookai.showError(" + JsString(FormatTurnError(ex)) + ");");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("outlookai.setComposerEnabled(true, false);");
                TraceLog.Write("<< StartTurnAsync", "InboxCopilot");
            }
        }

        private string BuildSystemInstructionsForCurrentState()
        {
            try
            {
                var sel = _surface?.GetCurrentSelection(includeFullBodies: false, maxItems: 1);
                int unreadCount = 0, totalCount = 0;
                string folderName = sel?.Folder ?? "Inbox";
                try
                {
                    var folder = _explorer?.CurrentFolder;
                    if (folder != null)
                    {
                        unreadCount = folder.UnReadItemCount;
                        totalCount = folder.Items.Count;
                    }
                }
                catch { }
                return InboxCopilotPromptBuilder.Build(folderName, unreadCount, totalCount, sel);
            }
            catch (Exception ex)
            {
                TraceLog.Write("BuildSystemInstructions error: " + ex, "InboxCopilot");
                return "You are the Outlook Inbox Copilot. Help the user with their mailbox.";
            }
        }

        private async Task StartCustomActionAsync(string actionId)
        {
            TraceLog.Write(">> StartCustomActionAsync id=" + actionId, "InboxCopilot");
            if (_turnInFlight || string.IsNullOrWhiteSpace(actionId) || !_isReady || _surface == null)
            {
                TraceLog.Write("StartCustomActionAsync aborted (gate)", "InboxCopilot");
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
                if (!string.IsNullOrWhiteSpace(result.DraftId))
                {
                    var draftInfo = new JObject(
                        new JProperty("id", result.DraftId),
                        new JProperty("title", result.DraftDisplayName ?? result.Text ?? ""),
                        new JProperty("location", result.DraftLocation ?? ""));
                    await RunScript("outlookai.onDraftCreated(" +
                        JsString(assistantId) + ", " + draftInfo.ToString(Newtonsoft.Json.Formatting.None) + ");");
                }
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", {});");
            }
            catch (OperationCanceledException)
            {
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", {stopped:true});");
            }
            catch (Exception ex)
            {
                TraceLog.Write("StartCustomActionAsync EXCEPTION: " + ex, "InboxCopilot");
                await RunScript("outlookai.showError(" + JsString(FormatTurnError(ex)) + ");");
                await RunScript("outlookai.finalizeAssistantMessage(" + JsString(assistantId) + ", {error:true});");
            }
            finally
            {
                _turnInFlight = false;
                _activeCts?.Dispose();
                _activeCts = null;
                await RunScript("outlookai.setComposerEnabled(true, false);");
                TraceLog.Write("<< StartCustomActionAsync", "InboxCopilot");
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
                var marshaller = Globals.ThisAddIn?.OutlookMarshaller;
                if (marshaller != null && Thread.CurrentThread.ManagedThreadId != marshaller.UiThreadId)
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
                var coreOnUi = _webView?.CoreWebView2;
                if (coreOnUi == null) return;
                await coreOnUi.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                TraceLog.Write("RunScript EXCEPTION: " + ex.Message, "InboxCopilot");
            }
        }

        private static string JsString(string s)
            => Newtonsoft.Json.JsonConvert.SerializeObject(s ?? "");

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            try { _activeCts?.Cancel(); } catch { }
            try { _recommendationCts?.Cancel(); } catch { }
            try { _recommendationCts?.Dispose(); } catch { }
            try
            {
                if (_explorer != null)
                {
                    _explorer.SelectionChange -= OnExplorerSelectionChange;
                    _explorer.FolderSwitch -= OnExplorerFolderSwitch;
                }
            }
            catch { }
            try { _webView?.Dispose(); } catch { }
        }

        private sealed class WebViewSink : ChatEventSink
        {
            private readonly InboxCopilotController _owner;
            private readonly string _assistantId;
            public bool HasReceivedText { get; private set; }
            public WebViewSink(InboxCopilotController owner, string assistantId)
            {
                _owner = owner;
                _assistantId = assistantId;
            }
            public override void OnTokenDelta(string delta)
            {
                if (!string.IsNullOrEmpty(delta)) HasReceivedText = true;
                _ = _owner.RunScript("outlookai.appendTextDelta(" +
                    JsString(_assistantId) + ", " + JsString(delta) + ");");
            }
            public override void OnToolCallStart(string callId, string name, string argsJson)
            {
                TraceLog.Write("Sink.OnToolCallStart " + name + " args=" + (argsJson?.Length > 200 ? argsJson.Substring(0, 200) + "..." : argsJson), "WebViewSink");
                _ = _owner.RunScript("outlookai.appendToolCallCard(" +
                    JsString(callId) + ", " + JsString(name) + ", " + JsString(argsJson) + ");");
            }
            public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson)
            {
                TraceLog.Write("Sink.OnToolCallResult ok=" + ok + " summary=" + summary + " resultLen=" + (resultJson?.Length ?? 0), "WebViewSink");
                _ = _owner.RunScript("outlookai.updateToolCallCard(" +
                    JsString(callId) + ", " + (ok ? "true" : "false") + ", " +
                    JsString(summary) + ", " + JsString(resultJson) + ");");
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
        }
    }
}
