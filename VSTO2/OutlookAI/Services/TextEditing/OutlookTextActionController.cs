using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.CustomActions;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Skills;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.Services.TextEditing
{
    /// <summary>
    /// Bridges Ribbon selection actions to Outlook's embedded Word editor.
    /// Word objects are intentionally late-bound so the add-in does not need
    /// a compile-time Microsoft.Office.Interop.Word reference.
    /// </summary>
    public sealed class OutlookTextActionController : IDisposable
    {
        private const uint GetAncestorRoot = 2;
        private const int SurroundingContextRadius = 6000;
        public const int MaximumSelectionCharacters = 32000;
        private const int MaximumErrorMessageLength = 500;
        private const string EmptySelectionMessage =
            "Выделите текст в теле элемента и повторите действие.";
        private const string StaleSelectionMessage =
            "Выделенный текст изменился, поэтому OutlookAI не стал его заменять.";

        private static readonly string EmptyMenuXml =
            TextActionMenuBuilder.Build(new CustomActionDefinition[0]);

        private readonly Outlook.Application _application;
        private readonly OutlookThreadMarshaller _marshaller;
        private readonly LiteLlmChatService _chat;
        private readonly CustomActionStore _customActionStore;
        private readonly TextActionHistoryStore _historyStore;
        private readonly object _gate = new object();
        private readonly Dictionary<long, ActiveOperation> _operations =
            new Dictionary<long, ActiveOperation>();
        private bool _disposed;

        public OutlookTextActionController(
            Outlook.Application application,
            OutlookThreadMarshaller marshaller,
            LiteLlmChatService chat,
            CustomActionStore customActionStore = null,
            TextActionHistoryStore historyStore = null)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _marshaller = marshaller ?? throw new ArgumentNullException(nameof(marshaller));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _customActionStore = customActionStore ?? new CustomActionStore();
            _historyStore = historyStore ?? new TextActionHistoryStore();
        }

        /// <summary>
        /// Returns true only for an editable Word body in a MailItem,
        /// AppointmentItem or TaskItem with a non-empty, safe selection.
        /// </summary>
        public bool CanExecute(Outlook.Inspector inspector)
        {
            if (IsDisposed) return false;
            try
            {
                return InvokeOnUi(() => CanExecuteCore(ResolveInspector(inspector)));
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Builds the Office dynamicMenu fragment from the current persisted
        /// text actions. Settings changes therefore become visible the next
        /// time the submenu is opened.
        /// </summary>
        public string GetMenuContent(Outlook.Inspector inspector)
        {
            if (IsDisposed) return EmptyMenuXml;
            try
            {
                return InvokeOnUi(() => GetMenuContentCore(ResolveInspector(inspector)));
            }
            catch
            {
                return EmptyMenuXml;
            }
        }

        /// <summary>
        /// Captures a duplicate Word Range before the first asynchronous model
        /// call, displays a modeless progress window and replaces the range
        /// only if it still contains the captured text.
        /// </summary>
        public async Task ExecuteAsync(Outlook.Inspector inspector, string actionId, string[] pinnedSkillIds = null)
        {
            if (string.IsNullOrWhiteSpace(actionId))
            {
                throw new ArgumentException("Text action id is required.", nameof(actionId));
            }
            ThrowIfDisposed();

            ActiveOperation operation;
            if (Thread.CurrentThread.ManagedThreadId == _marshaller.UiThreadId)
            {
                // Ribbon callbacks run here. This branch deliberately captures
                // Range.Duplicate synchronously before this method reaches its
                // first await.
                operation = PrepareExecution(ResolveInspector(inspector), actionId);
            }
            else
            {
                operation = await _marshaller.RunAsync(
                    () => PrepareExecution(ResolveInspector(inspector), actionId),
                    CancellationToken.None).ConfigureAwait(false);
            }

            if (operation == null) return;
            operation.PinnedSkillIds = pinnedSkillIds ?? new string[0];
            await RunPreparedAsync(operation).ConfigureAwait(false);
        }

        public void Dispose()
        {
            List<ActiveOperation> operations;
            lock (_gate)
            {
                if (_disposed) return;
                _disposed = true;
                operations = _operations.Values.Distinct().ToList();
            }

            Action closeAll = () =>
            {
                foreach (var operation in operations)
                {
                    operation.CancelSafely();
                    if (IsFormAlive(operation.Form))
                    {
                        operation.Form.Close();
                    }
                    else
                    {
                        CleanupOperation(operation);
                    }
                }
            };

            try
            {
                if (Thread.CurrentThread.ManagedThreadId == _marshaller.UiThreadId)
                {
                    closeAll();
                }
                else
                {
                    _marshaller.RunAsync(closeAll, CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
            }
            catch
            {
                foreach (var operation in operations) operation.CancelSafely();
            }
        }

        private bool CanExecuteCore(Outlook.Inspector inspector)
        {
            ItemContext itemContext;
            if (!TryGetItemContext(inspector, out itemContext)) return false;
            if (!HasUsableSelection(inspector)) return false;

            var actions = TextActionCatalog.GetActions(
                _customActionStore.LoadCatalog(),
                itemContext.ItemType,
                itemContext.Direction);
            return actions.Count > 0;
        }

        private string GetMenuContentCore(Outlook.Inspector inspector)
        {
            ItemContext itemContext;
            if (!TryGetItemContext(inspector, out itemContext)
                || !HasUsableSelection(inspector))
            {
                return EmptyMenuXml;
            }

            var actions = TextActionCatalog.GetActions(
                _customActionStore.LoadCatalog(),
                itemContext.ItemType,
                itemContext.Direction);
            return TextActionMenuBuilder.Build(actions);
        }

        private ActiveOperation PrepareExecution(
            Outlook.Inspector inspector,
            string actionId)
        {
            ThrowIfDisposed();

            ItemContext itemContext;
            if (!TryGetItemContext(inspector, out itemContext)) return null;

            var key = GetInspectorKey(inspector);
            ActiveOperation existing;
            lock (_gate)
            {
                _operations.TryGetValue(key, out existing);
            }
            if (existing != null && !existing.IsClosed)
            {
                FocusExisting(existing);
                return null;
            }

            var action = TextActionCatalog.GetActions(
                    _customActionStore.LoadCatalog(),
                    itemContext.ItemType,
                    itemContext.Direction)
                .FirstOrDefault(candidate => string.Equals(
                    candidate.Id,
                    actionId,
                    StringComparison.OrdinalIgnoreCase));
            if (action == null)
            {
                ShowMessage(
                    inspector,
                    "Это действие больше недоступно. Откройте меню ещё раз.");
                return null;
            }

            SelectionSnapshot snapshot;
            try
            {
                snapshot = CaptureSnapshot(inspector, itemContext);
            }
            catch (SelectionTooLongException ex)
            {
                ShowMessage(
                    inspector,
                    "Выделено слишком много текста ("
                        + ex.CharacterCount.ToString("N0")
                        + " знаков). Максимальный размер для одного AI-действия — "
                        + MaximumSelectionCharacters.ToString("N0")
                        + " знаков. Уменьшите выделение и повторите действие.");
                return null;
            }
            if (snapshot == null)
            {
                ShowMessage(inspector, EmptySelectionMessage);
                return null;
            }

            var form = new TextActionProgressForm();
            form.ShowRunning(action.Title);
            var operation = new ActiveOperation(
                this,
                key,
                inspector,
                action,
                snapshot,
                form);

            form.CancelRequested += (sender, args) => operation.CancelSafely();
            form.UndoRequested += async (sender, args) =>
                await UndoAsync(operation).ConfigureAwait(true);
            form.FormClosed += (sender, args) => CleanupOperation(operation);

            TrySubscribeInspectorClose(operation);
            var controllerWasDisposed = false;
            lock (_gate)
            {
                if (_disposed)
                {
                    controllerWasDisposed = true;
                }
                else
                {
                    _operations[key] = operation;
                }
            }
            if (controllerWasDisposed)
            {
                operation.TryMarkClosed();
                operation.CancelSafely();
                CleanupDetachedOperation(operation);
                operation.MarkRequestFinished();
                form.Dispose();
                throw new ObjectDisposedException(nameof(OutlookTextActionController));
            }

            try
            {
                if (snapshot.OwnerHandle != IntPtr.Zero)
                {
                    form.Show(new NativeWindowOwner(snapshot.OwnerHandle));
                }
                else
                {
                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.Show();
                }
            }
            catch
            {
                CleanupOperation(operation);
                form.Dispose();
                throw;
            }
            return operation;
        }

        private async Task RunPreparedAsync(ActiveOperation operation)
        {
            var skillSink = new TextActionSkillEventSink();
            try
            {
                var prompt = TextActionPromptBuilder.Build(
                    operation.Action,
                    operation.Snapshot.SelectedTextForModel,
                    operation.Snapshot.SurroundingContext);
                prompt.SystemPrompt += operation.Snapshot.Editor.Plan.LinkInstructions;

                string rawResult;
                if (operation.Action.UseSkills)
                {
                    if (!new SkillStore().Read().DisclosureAccepted)
                        throw new InvalidOperationException("Сначала откройте «Скиллы» в правой панели и подтвердите передачу знаний ИИ.");
                    var turn = await _chat.RunTurnAsync(new ConversationContext
                    {
                        SystemInstructions = prompt.SystemPrompt,
                        AllowedToolNames = new string[0],
                        IncludeWriteTools = false,
                        EnableSkills = true,
                        PinnedSkillIds = operation.PinnedSkillIds,
                        SuppressSkillAttribution = true
                    }, prompt.UserPrompt, new SkillSession(null, false), skillSink,
                        operation.Cancellation.Token).ConfigureAwait(false);
                    if (turn.StopReason != StopReason.Completed)
                        throw new InvalidOperationException("ИИ не завершил редактирование. Текст не изменён.");
                    rawResult = turn.FinalAssistantText;
                }
                else
                {
                    rawResult = await _chat.CompleteWithoutToolsAsync(
                        prompt.SystemPrompt, prompt.UserPrompt,
                        operation.Cancellation.Token).ConfigureAwait(false);
                }
                operation.Cancellation.Token.ThrowIfCancellationRequested();
                if (operation.IsClosed || operation.InspectorClosed) return;

                var safeResult = SanitizeModelOutput(rawResult);
                var replacement = TextActionPromptBuilder.NormalizeReplacement(
                    safeResult,
                    operation.Snapshot.Editor.Plan.ModelText);

                var applyResult = await _marshaller.RunAsync(
                    () => ApplyReplacement(operation, replacement),
                    CancellationToken.None).ConfigureAwait(false);

                if (applyResult.Closed) return;
                if (applyResult.Stale)
                {
                    await UpdateFormAsync(
                        operation,
                        form => form.ShowFailure(StaleSelectionMessage))
                        .ConfigureAwait(false);
                    return;
                }
                if (applyResult.NoChange)
                {
                    await UpdateFormAsync(operation, form => form.ShowNoChange())
                        .ConfigureAwait(false);
                    return;
                }

                operation.AppliedText = applyResult.AppliedText;
                await UpdateFormAsync(operation, form => form.ShowFinalizing())
                    .ConfigureAwait(false);

                TextActionHistoryEntry historyEntry = null;
                try
                {
                    historyEntry = await Task.Run(() => _historyStore.AppendApplied(
                        operation.Action.Id,
                        operation.Action.Title,
                        operation.Snapshot.ItemType,
                        operation.Snapshot.OriginalText,
                        applyResult.AppliedText)).ConfigureAwait(false);
                }
                catch
                {
                    // The edit has already succeeded. A history I/O failure
                    // must not roll it back or hide the Undo option.
                }
                operation.HistoryEntryId = historyEntry?.Id;

                await UpdateFormAsync(operation, form => form.ShowCompleted())
                    .ConfigureAwait(false);
                await _marshaller.RunAsync(() =>
                {
                    if (!operation.IsClosed && !operation.InspectorClosed)
                    {
                        operation.Inspector.Activate();
                        operation.Snapshot.Editor.Focus();
                    }
                    return true;
                }, CancellationToken.None).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                await UpdateFormAsync(operation, form => form.ShowCancelled())
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                if (!operation.IsClosed && !operation.InspectorClosed)
                {
                    var message = "Не удалось обработать выделенный текст. "
                        + GetSafeErrorMessage(ex);
                    await UpdateFormAsync(operation, form => form.ShowFailure(message))
                        .ConfigureAwait(false);
                }
            }
            finally
            {
                try
                {
                    // Metadata belongs in the status window, never in text inserted into Word.
                    // Updating this control does not activate the window or disturb the selection.
                    await UpdateFormAsync(operation, form => form.ShowLoadedSkills(skillSink.LoadedNames))
                        .ConfigureAwait(false);
                }
                finally { operation.MarkRequestFinished(); }
            }
        }

        private async Task UndoAsync(ActiveOperation operation)
        {
            if (!operation.TryBeginUndo()) return;
            try
            {
                await UpdateFormAsync(operation, form => form.ShowUndoInProgress())
                    .ConfigureAwait(false);

                var result = await _marshaller.RunAsync(
                    () => RestoreOriginal(operation),
                    CancellationToken.None).ConfigureAwait(false);
                if (result.Closed) return;
                if (!result.Restored)
                {
                    await UpdateFormAsync(
                        operation,
                        form => form.ShowUndoUnavailable(
                            "Текст после замены уже изменился. Чтобы не потерять новые правки, автоматическая отмена не выполнена."))
                        .ConfigureAwait(false);
                    return;
                }

                var historyUpdated = false;
                if (!string.IsNullOrWhiteSpace(operation.HistoryEntryId))
                {
                    try
                    {
                        historyUpdated = await Task.Run(() =>
                            _historyStore.MarkUndone(operation.HistoryEntryId))
                            .ConfigureAwait(false);
                    }
                    catch
                    {
                        historyUpdated = false;
                    }
                }

                await UpdateFormAsync(
                    operation,
                    form => form.ShowUndoCompleted(historyUpdated))
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                if (!operation.IsClosed && !operation.InspectorClosed)
                {
                    await UpdateFormAsync(
                        operation,
                        form => form.ShowUndoUnavailable(
                            "Не удалось отменить изменение. " + GetSafeErrorMessage(ex)))
                        .ConfigureAwait(false);
                }
            }
        }

        private ApplyResult ApplyReplacement(
            ActiveOperation operation,
            string replacement)
        {
            if (operation.IsClosed || operation.InspectorClosed)
            {
                return ApplyResult.ForClosed();
            }
            operation.Cancellation.Token.ThrowIfCancellationRequested();

            try
            {
                // Touch both objects on the Outlook UI thread. Either access
                // throws when the Inspector or its Word document was closed.
                var currentItem = operation.Inspector.CurrentItem;
                var currentEditor = operation.Inspector.WordEditor;
                if (currentItem == null || currentEditor == null)
                {
                    return ApplyResult.ForClosed();
                }
                if (!ReferenceEquals(currentEditor, operation.Snapshot.Document)) return ApplyResult.ForStale();

                dynamic range = operation.Snapshot.Range;
                var current = (string)range.Text ?? "";
                if (!string.Equals(
                    current,
                    operation.Snapshot.OriginalText,
                    StringComparison.Ordinal))
                {
                    return ApplyResult.ForStale();
                }
                if (!operation.Snapshot.Editor.IsCurrent) return ApplyResult.ForStale();
                var expected = operation.Snapshot.Editor.Plan.GetText(
                    operation.Snapshot.Editor.Plan.ParseReplacement(replacement));
                if (string.Equals(current, expected, StringComparison.Ordinal))
                {
                    return ApplyResult.ForNoChange();
                }

                // Cancellation can happen after the HTTP request completes but
                // before this UI delegate gets its turn in Outlook's queue.
                // Check again immediately before the only mutating COM call.
                operation.Cancellation.Token.ThrowIfCancellationRequested();
                var applied = operation.Snapshot.Editor.Apply(replacement,
                    "OutlookAI: " + (operation.Action.Title ?? operation.Action.Id));
                return ApplyResult.ForApplied(applied);
            }
            catch (COMException)
            {
                if (operation.IsClosed || operation.InspectorClosed)
                {
                    return ApplyResult.ForClosed();
                }
                throw;
            }
        }

        private UndoResult RestoreOriginal(ActiveOperation operation)
        {
            if (operation.IsClosed || operation.InspectorClosed)
            {
                return UndoResult.ForClosed();
            }

            try
            {
                var currentItem = operation.Inspector.CurrentItem;
                var currentEditor = operation.Inspector.WordEditor;
                if (currentItem == null || currentEditor == null)
                {
                    return UndoResult.ForClosed();
                }
                if (!ReferenceEquals(currentEditor, operation.Snapshot.Document)) return UndoResult.ForStale();

                dynamic range = operation.Snapshot.Range;
                var current = (string)range.Text ?? "";
                if (string.Equals(
                    current,
                    operation.Snapshot.OriginalText,
                    StringComparison.Ordinal))
                {
                    // The user already used Word's Ctrl+Z while the status
                    // window was open. Do not write the original a second
                    // time, but do reconcile the persistent history record.
                    return UndoResult.ForRestored();
                }
                if (string.IsNullOrEmpty(operation.AppliedText)
                    || !string.Equals(
                        current,
                        operation.AppliedText,
                        StringComparison.Ordinal))
                {
                    return UndoResult.ForStale();
                }

                return operation.Snapshot.Editor.Undo()
                    ? UndoResult.ForRestored() : UndoResult.ForStale();
            }
            catch (COMException)
            {
                if (operation.IsClosed || operation.InspectorClosed)
                {
                    return UndoResult.ForClosed();
                }
                throw;
            }
        }

        private SelectionSnapshot CaptureSnapshot(
            Outlook.Inspector inspector,
            ItemContext itemContext)
        {
            dynamic duplicate = null;
            object activeWindow = null;
            object selection = null;
            object sourceRange = null;
            WordSelectionEditor editor = null;
            try
            {
                if (!HasWordEditor(inspector)) return null;
                dynamic document = inspector.WordEditor;
                if (!IsDocumentEditable(document)) return null;

                activeWindow = document.ActiveWindow;
                selection = ((dynamic)activeWindow).Selection;
                sourceRange = ((dynamic)selection).Range;
                duplicate = ((dynamic)sourceRange).Duplicate;
                var original = AdjustAndReadSelection(duplicate);
                if (!IsSafeSelection(original)) return null;
                if (original.Length > MaximumSelectionCharacters)
                {
                    throw new SelectionTooLongException(original.Length);
                }

                editor = WordSelectionEditor.Capture((object)document, (object)duplicate);
                var selectedForModel = SanitizeWordTextForModel(editor.Plan.ModelText);
                if (string.IsNullOrWhiteSpace(selectedForModel)) return null;

                var start = (int)duplicate.Start;
                var end = (int)duplicate.End;
                var surrounding = CaptureSurroundingContext(
                    document,
                    start,
                    end,
                    itemContext);
                var handle = GetWordWindowHandle(activeWindow);

                var snapshot = new SelectionSnapshot
                {
                    Document = document,
                    Range = duplicate,
                    Editor = editor,
                    OriginalText = original,
                    SelectedTextForModel = selectedForModel,
                    SurroundingContext = surrounding,
                    ItemType = itemContext.ItemType,
                    OwnerHandle = handle
                };
                duplicate = null;
                editor = null;
                return snapshot;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                editor?.Dispose();
                TryReleaseComObject(duplicate);
                TryReleaseComObject(sourceRange);
                TryReleaseComObject(selection);
                TryReleaseComObject(activeWindow);
            }
        }

        private bool HasUsableSelection(Outlook.Inspector inspector)
        {
            dynamic duplicate = null;
            object activeWindow = null;
            object selection = null;
            object sourceRange = null;
            try
            {
                if (!HasWordEditor(inspector)) return false;
                dynamic document = inspector.WordEditor;
                if (!IsDocumentEditable(document)) return false;
                activeWindow = document.ActiveWindow;
                selection = ((dynamic)activeWindow).Selection;
                sourceRange = ((dynamic)selection).Range;
                duplicate = ((dynamic)sourceRange).Duplicate;
                var original = AdjustAndReadSelection(duplicate);
                return IsSafeSelection(original)
                    && !string.IsNullOrWhiteSpace(
                        SanitizeWordTextForModel(original));
            }
            catch
            {
                return false;
            }
            finally
            {
                TryReleaseComObject(duplicate);
                TryReleaseComObject(sourceRange);
                TryReleaseComObject(selection);
                TryReleaseComObject(activeWindow);
            }
        }

        internal static string AdjustAndReadSelection(object rangeObject)
        {
            dynamic range = rangeObject;
            var start = (int)range.Start;
            var end = (int)range.End;
            if (end <= start) return "";

            var text = (string)range.Text ?? "";
            // A Word table-cell selection ends with paragraph + end-of-cell
            // markers in Range.Text, but both occupy ONE Word range position.
            // Never replace them or subtract two positions (which drops the
            // final real character from the selection).
            if (text.EndsWith("\r\a", StringComparison.Ordinal)
                && end > start)
            {
                range.End = end - 1;
                text = (string)range.Text ?? "";
            }
            return text;
        }

        private static bool IsSafeSelection(string original)
        {
            if (string.IsNullOrEmpty(original)) return false;
            // An internal end-of-cell marker means that the range spans cells.
            // Replacing such a range could change the table structure.
            return original.IndexOf('\a') < 0;
        }

        private static string CaptureSurroundingContext(
            dynamic document,
            int selectionStart,
            int selectionEnd,
            ItemContext itemContext)
        {
            dynamic content = null;
            dynamic beforeRange = null;
            dynamic afterRange = null;
            try
            {
                content = document.Content;
                var documentStart = (int)content.Start;
                var documentEnd = (int)content.End;
                var beforeStart = Math.Max(
                    documentStart,
                    selectionStart - SurroundingContextRadius);
                var afterEnd = Math.Min(
                    documentEnd,
                    selectionEnd + SurroundingContextRadius);

                var before = "";
                var after = "";
                if (beforeStart < selectionStart)
                {
                    beforeRange = document.Range(beforeStart, selectionStart);
                    before = SanitizeWordTextForModel((string)beforeRange.Text ?? "");
                }
                if (selectionEnd < afterEnd)
                {
                    afterRange = document.Range(selectionEnd, afterEnd);
                    after = SanitizeWordTextForModel((string)afterRange.Text ?? "");
                }

                var context = new StringBuilder();
                context.AppendLine("Тип элемента Outlook: " + itemContext.ItemType);
                if (!string.IsNullOrWhiteSpace(itemContext.Subject))
                {
                    context.AppendLine("Тема: "
                        + SanitizeWordTextForModel(itemContext.Subject));
                }
                if (!string.IsNullOrWhiteSpace(before))
                {
                    context.AppendLine("Текст перед выделением:");
                    context.AppendLine(before);
                }
                if (!string.IsNullOrWhiteSpace(after))
                {
                    context.AppendLine("Текст после выделения:");
                    context.AppendLine(after);
                }
                return context.ToString().Trim();
            }
            catch
            {
                return "Тип элемента Outlook: " + itemContext.ItemType;
            }
            finally
            {
                TryReleaseComObject(afterRange);
                TryReleaseComObject(beforeRange);
                TryReleaseComObject(content);
            }
        }

        private static bool TryGetItemContext(
            Outlook.Inspector inspector,
            out ItemContext context)
        {
            context = null;
            if (inspector == null) return false;
            try
            {
                if (inspector.EditorType != Outlook.OlEditorType.olEditorWord)
                {
                    return false;
                }

                var item = inspector.CurrentItem;
                var mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    if (mail.Sent) return false;
                    context = new ItemContext
                    {
                        ItemType = "mail",
                        Direction = "outgoing",
                        Subject = mail.Subject ?? ""
                    };
                    return true;
                }

                var appointment = item as Outlook.AppointmentItem;
                if (appointment != null)
                {
                    context = new ItemContext
                    {
                        ItemType = "meeting",
                        Direction = "outgoing",
                        Subject = appointment.Subject ?? ""
                    };
                    return true;
                }

                var task = item as Outlook.TaskItem;
                if (task != null)
                {
                    context = new ItemContext
                    {
                        ItemType = "task",
                        Direction = "outgoing",
                        Subject = task.Subject ?? ""
                    };
                    return true;
                }
                return false;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static bool HasWordEditor(Outlook.Inspector inspector)
        {
            if (inspector == null) return false;
            object activeWindow = null;
            try
            {
                // Do not require Inspector.IsWordMail(): that check is
                // mail-specific and may be false for appointment/task editors.
                if (inspector.EditorType != Outlook.OlEditorType.olEditorWord)
                {
                    return false;
                }
                dynamic document = inspector.WordEditor;
                if (document == null) return false;
                activeWindow = document.ActiveWindow;
                return activeWindow != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                TryReleaseComObject(activeWindow);
            }
        }

        private static bool IsDocumentEditable(dynamic document)
        {
            try
            {
                return document != null && !(bool)document.ReadOnly;
            }
            catch
            {
                // Some embedded Word document versions do not expose ReadOnly
                // through IDispatch. Successful range access is then the best
                // available editability check.
                return document != null;
            }
        }

        private static string SanitizeWordTextForModel(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            var safe = new StringBuilder(value.Length);
            for (var i = 0; i < value.Length; i++)
            {
                var ch = value[i];
                if (ch == '\a') continue;
                if (ch == '\r' || ch == '\v' || ch == '\f')
                {
                    safe.Append('\n');
                    continue;
                }
                if (ch == '\n' || ch == '\t')
                {
                    safe.Append(ch);
                    continue;
                }
                if (char.IsControl(ch)) continue;
                safe.Append(ch);
            }
            return safe.ToString();
        }

        private static string SanitizeModelOutput(string value)
        {
            if (value == null) return null;
            var safe = new StringBuilder(value.Length);
            for (var i = 0; i < value.Length; i++)
            {
                var ch = value[i];
                if (ch == '\r' || ch == '\n' || ch == '\t')
                {
                    safe.Append(ch);
                    continue;
                }
                if (char.IsControl(ch)) continue;
                safe.Append(ch);
            }
            return safe.ToString();
        }

        private void TrySubscribeInspectorClose(ActiveOperation operation)
        {
            try
            {
                operation.InspectorEvents =
                    (Outlook.InspectorEvents_10_Event)operation.Inspector;
                operation.InspectorEvents.Close += operation.HandleInspectorClosed;
            }
            catch
            {
                operation.InspectorEvents = null;
            }
        }

        private void OnInspectorClosed(ActiveOperation operation)
        {
            operation.InspectorClosed = true;
            operation.CancelSafely();
            Action close = () =>
            {
                if (IsFormAlive(operation.Form)) operation.Form.Close();
                else CleanupOperation(operation);
            };
            try
            {
                if (Thread.CurrentThread.ManagedThreadId == _marshaller.UiThreadId)
                {
                    close();
                }
                else
                {
                    _marshaller.RunAsync(close, CancellationToken.None);
                }
            }
            catch
            {
            }
        }

        private Task UpdateFormAsync(
            ActiveOperation operation,
            Action<TextActionProgressForm> update)
        {
            try
            {
                return _marshaller.RunAsync(() =>
                {
                    if (!operation.IsClosed && IsFormAlive(operation.Form))
                    {
                        update(operation.Form);
                    }
                }, CancellationToken.None);
            }
            catch
            {
                return Task.CompletedTask;
            }
        }

        private void CleanupOperation(ActiveOperation operation)
        {
            if (operation == null || !operation.TryMarkClosed()) return;
            operation.CancelSafely();
            lock (_gate)
            {
                ActiveOperation current;
                if (_operations.TryGetValue(operation.Key, out current)
                    && ReferenceEquals(current, operation))
                {
                    _operations.Remove(operation.Key);
                }
            }

            CleanupDetachedOperation(operation);
            operation.TryDisposeCancellation();
        }

        private static void CleanupDetachedOperation(ActiveOperation operation)
        {
            if (operation.InspectorEvents != null)
            {
                try
                {
                    operation.InspectorEvents.Close -= operation.HandleInspectorClosed;
                }
                catch
                {
                }
                operation.InspectorEvents = null;
            }
            TryReleaseComObject(operation.Snapshot.Range);
            operation.Snapshot.Editor?.Dispose();
            operation.Snapshot.Editor = null;
            operation.Snapshot.Range = null;
            operation.Snapshot.Document = null;
        }

        private static void FocusExisting(ActiveOperation operation)
        {
            try
            {
                if (!IsFormAlive(operation.Form)) return;
                if (operation.Form.WindowState == FormWindowState.Minimized)
                {
                    operation.Form.WindowState = FormWindowState.Normal;
                }
                operation.Form.Show();
                operation.Form.Activate();
                operation.Form.BringToFront();
            }
            catch
            {
            }
        }

        private Outlook.Inspector ResolveInspector(Outlook.Inspector inspector)
        {
            return inspector ?? _application.ActiveInspector();
        }

        private T InvokeOnUi<T>(Func<T> callback)
        {
            if (Thread.CurrentThread.ManagedThreadId == _marshaller.UiThreadId)
            {
                return callback();
            }
            return _marshaller.RunAsync(callback, CancellationToken.None)
                .GetAwaiter().GetResult();
        }

        private static long GetInspectorKey(Outlook.Inspector inspector)
        {
            var handle = GetInspectorHandle(inspector);
            if (handle != IntPtr.Zero) return handle.ToInt64();
            return long.MinValue + RuntimeHelpers.GetHashCode(inspector);
        }

        private static IntPtr GetInspectorHandle(Outlook.Inspector inspector)
        {
            object activeWindow = null;
            try
            {
                if (inspector == null) return IntPtr.Zero;
                dynamic document = inspector.WordEditor;
                if (document == null) return IntPtr.Zero;
                activeWindow = document.ActiveWindow;
                return GetWordWindowHandle(activeWindow);
            }
            catch
            {
                return IntPtr.Zero;
            }
            finally
            {
                TryReleaseComObject(activeWindow);
            }
        }

        private static IntPtr GetWordWindowHandle(object activeWindow)
        {
            if (activeWindow == null) return IntPtr.Zero;
            try
            {
                var hwnd = (int)((dynamic)activeWindow).Hwnd;
                if (hwnd == 0) return IntPtr.Zero;
                // Office exposes HWND as a signed 32-bit COM Long even in a
                // 64-bit process. Zero-extension avoids corrupting handles
                // whose high 32-bit sign bit is set.
                var window = new IntPtr(unchecked((long)(uint)hwnd));
                var root = GetAncestor(window, GetAncestorRoot);
                return root == IntPtr.Zero ? window : root;
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

        private static void ShowMessage(Outlook.Inspector inspector, string message)
        {
            var handle = GetInspectorHandle(inspector);
            if (handle == IntPtr.Zero)
            {
                MessageBox.Show(
                    message,
                    "OutlookAI",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }
            MessageBox.Show(
                new NativeWindowOwner(handle),
                message,
                "OutlookAI",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static bool IsFormAlive(Form form)
        {
            return form != null && !form.IsDisposed && !form.Disposing;
        }

        private static void TryReleaseComObject(object value)
        {
            if (value == null) return;
            try
            {
                if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
            }
            catch
            {
            }
        }

        private static string GetSafeErrorMessage(Exception exception)
        {
            var current = exception;
            while (current is AggregateException && current.InnerException != null)
            {
                current = current.InnerException;
            }
            var message = (current?.Message ?? "").Trim();
            foreach (var separator in new[] { '\r', '\n', '\0' })
            {
                message = message.Replace(separator, ' ');
            }
            if (message.Length > MaximumErrorMessageLength)
            {
                message = message.Substring(0, MaximumErrorMessageLength) + "…";
            }
            return message;
        }

        private bool IsDisposed
        {
            get
            {
                lock (_gate) return _disposed;
            }
        }

        private void ThrowIfDisposed()
        {
            if (IsDisposed)
            {
                throw new ObjectDisposedException(nameof(OutlookTextActionController));
            }
        }

        private sealed class NativeWindowOwner : IWin32Window
        {
            public NativeWindowOwner(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
        }

        private sealed class ItemContext
        {
            public string ItemType { get; set; }
            public string Direction { get; set; }
            public string Subject { get; set; }
        }

        private sealed class SelectionSnapshot
        {
            public object Document { get; set; }
            public object Range { get; set; }
            public WordSelectionEditor Editor { get; set; }
            public string OriginalText { get; set; }
            public string SelectedTextForModel { get; set; }
            public string SurroundingContext { get; set; }
            public string ItemType { get; set; }
            public IntPtr OwnerHandle { get; set; }
        }

        private sealed class SelectionTooLongException : Exception
        {
            public SelectionTooLongException(int characterCount)
            {
                CharacterCount = characterCount;
            }

            public int CharacterCount { get; }
        }

        private sealed class ActiveOperation
        {
            private readonly object _stateGate = new object();
            private bool _closed;
            private bool _undoStarted;
            private bool _requestFinished;
            private bool _cancellationDisposed;

            public ActiveOperation(
                OutlookTextActionController owner,
                long key,
                Outlook.Inspector inspector,
                CustomActionDefinition action,
                SelectionSnapshot snapshot,
                TextActionProgressForm form)
            {
                Owner = owner;
                Key = key;
                Inspector = inspector;
                Action = action;
                Snapshot = snapshot;
                Form = form;
                Cancellation = new CancellationTokenSource();
            }

            public OutlookTextActionController Owner { get; }
            public long Key { get; }
            public Outlook.Inspector Inspector { get; }
            public CustomActionDefinition Action { get; }
            public SelectionSnapshot Snapshot { get; }
            public TextActionProgressForm Form { get; }
            public CancellationTokenSource Cancellation { get; }
            public Outlook.InspectorEvents_10_Event InspectorEvents { get; set; }
            public volatile bool InspectorClosed;
            public string AppliedText { get; set; }
            public string HistoryEntryId { get; set; }
            public string[] PinnedSkillIds { get; set; }

            public bool IsClosed
            {
                get
                {
                    lock (_stateGate) return _closed;
                }
            }

            public void HandleInspectorClosed()
            {
                Owner.OnInspectorClosed(this);
            }

            public bool TryMarkClosed()
            {
                lock (_stateGate)
                {
                    if (_closed) return false;
                    _closed = true;
                    return true;
                }
            }

            public bool TryBeginUndo()
            {
                lock (_stateGate)
                {
                    if (_closed || _undoStarted) return false;
                    _undoStarted = true;
                    return true;
                }
            }

            public void CancelSafely()
            {
                lock (_stateGate)
                {
                    if (_cancellationDisposed) return;
                    try { Cancellation.Cancel(); }
                    catch (ObjectDisposedException) { }
                }
            }

            public void MarkRequestFinished()
            {
                lock (_stateGate) _requestFinished = true;
                TryDisposeCancellation();
            }

            public void TryDisposeCancellation()
            {
                lock (_stateGate)
                {
                    if (_cancellationDisposed || !_requestFinished || !_closed) return;
                    Cancellation.Dispose();
                    _cancellationDisposed = true;
                }
            }
        }

        private sealed class ApplyResult
        {
            public bool Closed { get; private set; }
            public bool Stale { get; private set; }
            public bool NoChange { get; private set; }
            public string AppliedText { get; private set; }

            public static ApplyResult ForClosed()
            {
                return new ApplyResult { Closed = true };
            }

            public static ApplyResult ForStale()
            {
                return new ApplyResult { Stale = true };
            }

            public static ApplyResult ForNoChange()
            {
                return new ApplyResult { NoChange = true };
            }

            public static ApplyResult ForApplied(string value)
            {
                return new ApplyResult { AppliedText = value ?? "" };
            }
        }

        private sealed class UndoResult
        {
            public bool Closed { get; private set; }
            public bool Restored { get; private set; }

            public static UndoResult ForClosed()
            {
                return new UndoResult { Closed = true };
            }

            public static UndoResult ForStale()
            {
                return new UndoResult();
            }

            public static UndoResult ForRestored()
            {
                return new UndoResult { Restored = true };
            }
        }
    }

    /// <summary>Collects only successful skill loads for the selected-text action's status UI.</summary>
    public sealed class TextActionSkillEventSink : ChatEventSink
    {
        private readonly object _gate = new object();
        private readonly HashSet<string> _loadCalls = new HashSet<string>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _loaded = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public IReadOnlyList<string> LoadedNames
        {
            get
            {
                lock (_gate) return _loaded.Values.Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(name => name, StringComparer.CurrentCulture).ToArray();
            }
        }

        public override void OnToolCallStart(string callId, string name, string argsJson)
        {
            if (string.IsNullOrEmpty(callId) || !string.Equals(name, SkillToolNames.Load, StringComparison.OrdinalIgnoreCase)) return;
            lock (_gate) _loadCalls.Add(callId);
        }

        public override void OnToolCallResult(string callId, bool ok, string summary, string resultJson)
        {
            lock (_gate)
            {
                if (callId == null || !_loadCalls.Remove(callId) || !ok) return;
                try
                {
                    var result = JObject.Parse(resultJson ?? "{}");
                    if (result["error"] != null || !(result["skills"] is JArray skills)) return;
                    foreach (var skill in skills.OfType<JObject>())
                    {
                        var id = (string)skill["id"];
                        var name = ((string)skill["name"] ?? "").Replace('\r', ' ').Replace('\n', ' ').Trim();
                        if (!string.IsNullOrWhiteSpace(id) && name.Length > 0) _loaded[id] = name;
                    }
                }
                catch (Newtonsoft.Json.JsonException) { /* Status metadata must not fail the edit. */ }
            }
        }
    }
}
