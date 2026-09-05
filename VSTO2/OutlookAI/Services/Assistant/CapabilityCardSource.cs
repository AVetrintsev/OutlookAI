using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.CustomActions;
using OutlookAI.Services.Tools;

namespace OutlookAI.Services.Assistant
{
    public sealed class CapabilityCardSource
    {
        private readonly CustomActionStore _customActionStore;

        public CapabilityCardSource(CustomActionStore customActionStore = null)
        {
            _customActionStore = customActionStore ?? new CustomActionStore();
        }

        public CapabilityCard[] LoadCards(bool includeWriteTools)
        {
            var cards = new List<CapabilityCard>();
            cards.AddRange(LoadToolCards(includeWriteTools));
            cards.AddRange(LoadCustomActionCards());
            return cards
                .Where(card => !string.IsNullOrWhiteSpace(card.Id))
                .GroupBy(card => card.Id, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToArray();
        }

        private static IEnumerable<CapabilityCard> LoadToolCards(bool includeWriteTools)
        {
            foreach (var tool in ToolManifestCatalog.Default.BuildUiToolsArray().OfType<JObject>())
            {
                var name = (string)tool["name"] ?? "";
                if (string.IsNullOrWhiteSpace(name)) continue;
                var isWrite = (bool?)tool["is_write"] ?? ToolManifestCatalog.IsWriteTool(name);
                if (isWrite && !includeWriteTools) continue;
                yield return new CapabilityCard
                {
                    Id = "tool:" + name,
                    Type = "tool",
                    Group = GroupForTool(name),
                    Title = name,
                    Description = (string)tool["description"] ?? name,
                    WhenToUse = WhenToUseForTool(name),
                    WhenNotToUse = new[] { "when selected metadata is sufficient" },
                    RequiredCapabilities = new[] { isWrite ? "mail.write" : "mail.read" },
                    Risk = isWrite ? "write" : RiskForTool(name),
                    ToolNames = new[] { name }
                };
            }
        }

        private IEnumerable<CapabilityCard> LoadCustomActionCards()
        {
            CustomActionFile catalog;
            try { catalog = _customActionStore.LoadCatalog(); }
            catch { yield break; }

            foreach (var action in (catalog.Groups ?? new CustomActionGroup[0])
                .SelectMany(group => group.Actions ?? new CustomActionDefinition[0]))
            {
                if (action == null
                    || action.Disabled
                    || CustomActionSurface.Normalize(action.Surface) != CustomActionSurface.Assistant)
                {
                    continue;
                }
                var allowedTools = action.AllowedTools ?? new string[0];
                yield return new CapabilityCard
                {
                    Id = "action:" + (action.Id ?? action.Title ?? Guid.NewGuid().ToString("N")),
                    Type = "action",
                    Group = "custom.action",
                    Title = action.Title ?? "",
                    Description = action.Description ?? action.Prompt ?? "",
                    WhenToUse = new[] { action.Title ?? "", action.Description ?? "", action.Prompt ?? "" },
                    RequiredCapabilities = allowedTools.Any(ToolManifestCatalog.IsWriteTool)
                        ? new[] { "mail.read", "mail.write" }
                        : new[] { "mail.read" },
                    Risk = allowedTools.Any(ToolManifestCatalog.IsWriteTool) ? "write" : "read_only",
                    ToolNames = allowedTools
                };
            }
        }

        private static string GroupForTool(string name)
        {
            if (name == "outlook_get_current_selection"
                || name == "outlook_read_message"
                || name == "outlook_read_conversation"
                || name == "outlook_read_messages"
                || name == "outlook_get_current_compose_state")
            {
                return "selection.read";
            }
            if (name == "outlook_search_messages"
                || name == "outlook_count_messages"
                || name == "outlook_aggregate_messages"
                || name == "outlook_list_folders")
            {
                return "mail.search";
            }
            if (name == "outlook_export_excel"
                || name == "outlook_export_pdf"
                || name == "outlook_export_search_results")
            {
                return "export";
            }
            if (ToolManifestCatalog.IsWriteTool(name))
            {
                return "mail.write";
            }
            return "tool";
        }

        private static string RiskForTool(string name)
        {
            return GroupForTool(name) == "export" ? "export" : "read_only";
        }

        private static string[] WhenToUseForTool(string name)
        {
            switch (GroupForTool(name))
            {
                case "selection.read":
                    return new[] { "selected message", "current email", "body of selected item", "who sent this", "why was this sent", "выбранное письмо", "текущее письмо", "кто отправил", "зачем прислали", "почему отправили" };
                case "mail.search":
                    return new[] { "find messages", "count messages", "search mailbox", "aggregate emails", "emails from sender", "найди письма", "поиск писем", "сколько писем", "письма от отправителя" };
                case "export":
                    return new[] { "create report file", "excel export", "pdf export", "spreadsheet", "отчет", "экспорт", "excel", "pdf", "таблица" };
                case "mail.write":
                    return new[] { "create draft", "mark read", "flag message", "set category", "создать черновик", "ответить", "пометить", "категория" };
                default:
                    return new[] { name };
            }
        }
    }
}
