using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services.Skills
{
    public sealed class SkillUiBridge : IDisposable
    {
        private readonly SkillStore _store;
        private readonly SkillDraftService _drafts;
        private readonly IOutlookSurface _surface;
        private readonly CancellationTokenSource _lifetime = new CancellationTokenSource();
        private CancellationTokenSource _draftCancellation;
        private SkillDefinition[] _pendingImport;
        private long _importRevision;
        private string _importToken;
        private int _draftBusy;
        private string[] _pinnedIds = new string[0];

        public string[] PinnedIds => _pinnedIds.ToArray();

        public SkillUiBridge(LiteLlmChatService chat, IOutlookSurface surface, SkillStore store = null)
        {
            _store = store ?? new SkillStore();
            _drafts = new SkillDraftService(chat);
            _surface = surface;
        }

        public Task PushAsync(Func<string, Task> send)
        {
            try
            {
                var file = _store.Read();
                _pinnedIds = _pinnedIds.Where(id => file.Skills.Any(skill => skill.Id == id && skill.IsActive(DateTime.Today))).ToArray();
                return Send(send, new { type = "catalog", skills = file.Skills,
                    disclosure_accepted = file.DisclosureAccepted, pinned_ids = _pinnedIds });
            }
            catch (Exception ex) { return Send(send, new { type = "error", message = "Не удалось прочитать скиллы: " + ex.Message }); }
        }

        public async Task<bool> HandleAsync(string type, JObject payload, Func<string, Task> send)
        {
            if (!(type ?? "").StartsWith("skills_", StringComparison.Ordinal)) return false;
            payload = payload ?? new JObject();
            try
            {
                switch (type)
                {
                    case "skills_request": await PushAsync(send); break;
                    case "skills_disclosure": _store.AcceptDisclosure(); await PushAsync(send); break;
                    case "skills_save":
                        var skill = SkillJson.Parse<SkillDefinition>(payload["skill"]?.ToString() ?? "null");
                        var saved = _store.Upsert(skill, (int?)payload["expected_revision"] ?? 0);
                        await Send(send, new { type = "saved", id = saved.Id });
                        await PushAsync(send);
                        break;
                    case "skills_delete":
                        _store.Delete((string)payload["id"], (int?)payload["revision"] ?? -1);
                        await PushAsync(send);
                        break;
                    case "skills_pin":
                        var ids = payload["ids"]?.ToObject<string[]>() ?? new string[0];
                        if (ids.Length > 5) throw new ArgumentException("Можно закрепить до 5 скиллов для этой панели.");
                        _pinnedIds = ids.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                        await PushAsync(send);
                        break;
                    case "skills_export":
                        await Send(send, new { type = "export", json = _store.Export() }); break;
                    case "skills_import_preview":
                        _pendingImport = _store.ParseImport((string)payload["json"]);
                        var file = _store.Read();
                        _importRevision = file.Revision;
                        _importToken = Guid.NewGuid().ToString("N");
                        await Send(send, new { type = "import_preview", token = _importToken,
                            items = _pendingImport.Select(item => new { item.Id, item.Name,
                                conflict = file.Skills.Any(current => string.Equals(current.Id, item.Id, StringComparison.OrdinalIgnoreCase)) }) });
                        break;
                    case "skills_import_apply":
                        if (_pendingImport == null || (string)payload["token"] != _importToken)
                            throw new InvalidOperationException("Сначала проверьте предварительный просмотр импорта.");
                        _store.ApplyImport(_pendingImport, _importRevision, payload["replace_ids"]?.ToObject<string[]>());
                        _pendingImport = null;
                        _importToken = null;
                        await Send(send, new { type = "imported" });
                        await PushAsync(send);
                        break;
                    case "skills_ai_cancel": _draftCancellation?.Cancel(); break;
                    case "skills_ai_draft": await CreateDraftAsync(payload, send); break;
                }
            }
            catch (OperationCanceledException) { await Send(send, new { type = "ai_cancelled", request_id = (string)payload["request_id"] }); }
            catch (Exception ex) { await Send(send, new { type = "error", message = ex.Message, request_id = (string)payload["request_id"] }); }
            return true;
        }

        private async Task CreateDraftAsync(JObject payload, Func<string, Task> send)
        {
            var requestId = (string)payload["request_id"];
            if (Interlocked.CompareExchange(ref _draftBusy, 1, 0) != 0)
                throw new InvalidOperationException("Дождитесь текущего черновика или отмените его.");
            try
            {
                var file = _store.Read();
                if (!file.DisclosureAccepted) throw new InvalidOperationException("Прочитайте уведомление о передаче знаний модели и нажмите «Понятно».");
                var id = (string)payload["id"];
                var existing = file.Skills.FirstOrDefault(skill => string.Equals(skill.Id, id, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrWhiteSpace(id) && existing == null) throw new InvalidOperationException("Скилл для обновления не найден.");
                _draftCancellation = CancellationTokenSource.CreateLinkedTokenSource(_lifetime.Token);
                await Send(send, new { type = "ai_busy", busy = true, request_id = requestId });
                var ct = _draftCancellation.Token;
                ConversationResult context = null;
                if ((bool?)payload["include_context"] == true)
                    context = await Task.Run(() => ConversationResult.ReadCurrent(_surface, 100, ct), ct).ConfigureAwait(false);
                var draft = await _drafts.CreateAsync((string)payload["description"], context, existing, ct).ConfigureAwait(false);
                ct.ThrowIfCancellationRequested();
                await Send(send, new { type = "draft", draft, request_id = requestId });
            }
            finally
            {
                _draftCancellation?.Dispose();
                _draftCancellation = null;
                Interlocked.Exchange(ref _draftBusy, 0);
                await Send(send, new { type = "ai_busy", busy = false, request_id = requestId });
            }
        }

        private static Task Send(Func<string, Task> send, object payload) => send(
            "window.outlookSkills.receive(" + SkillJson.Write(payload) + ");");

        public void Dispose()
        {
            _lifetime.Cancel();
            _lifetime.Dispose();
        }
    }
}
