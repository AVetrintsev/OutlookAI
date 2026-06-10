using System.Collections.Generic;

namespace OutlookAI.TaskPane.InboxCopilot
{
    /// <summary>
    /// A pre-canned prompt rendered as a clickable chip above the Inbox
    /// Copilot composer. Clicking a chip pre-fills the textarea and
    /// auto-sends (per Phase 3a spec). The set is computed server-side
    /// based on how many messages the user has selected in the active
    /// Explorer.
    /// </summary>
    public sealed class QuickActionChip
    {
        public string Label { get; set; }    // shown on the button
        public string Prompt { get; set; }   // pre-filled into the textarea

        /// <summary>
        /// Build the default chip set for a given selection count.
        /// Three static chips plus 0 or 2 dynamic chips:
        ///   0 selected -> static only
        ///   1 selected -> static + summarize thread + draft reply
        ///   2+ selected -> static + summarize selected + triage selected
        /// </summary>
        public static IReadOnlyList<QuickActionChip> ComputeChipsForSelectionCount(int selectionCount)
        {
            var list = new List<QuickActionChip>
            {
                new QuickActionChip
                {
                    Label = "Что требует внимания?",
                    Prompt = "Посмотри мой почтовый ящик и скажи, что требует внимания. Расставь приоритеты по свежести, важности и отправителю. Ответь кратко.",
                },
                new QuickActionChip
                {
                    Label = "Сводка непрочитанных",
                    Prompt = "Сделай сводку всех непрочитанных сообщений. Сгруппируй по отправителю или теме. Ответь кратко.",
                },
                new QuickActionChip
                {
                    Label = "Письма за сегодня",
                    Prompt = "Покажи всё, что я получил сегодня, сгруппируй по отправителю. Выдели всё, что выглядит срочным.",
                },
            };

            if (selectionCount == 1)
            {
                list.Add(new QuickActionChip
                {
                    Label = "Сводка переписки",
                    Prompt = "Сделай сводку выбранного сообщения и всей связанной переписки.",
                });
                list.Add(new QuickActionChip
                {
                    Label = "Черновик ответа",
                    Prompt = "Подготовь черновик ответа на выбранное сообщение. Сохрани тон отправителя.",
                });
            }
            else if (selectionCount >= 2)
            {
                list.Add(new QuickActionChip
                {
                    Label = "Сводка выбранных",
                    Prompt = "Сделай сводку всех выбранных сообщений.",
                });
                list.Add(new QuickActionChip
                {
                    Label = "Разобрать выбранные",
                    Prompt = "Разбери выбранные сообщения: какие требуют действия, какие можно архивировать, какие можно отметить прочитанными?",
                });
            }

            return list;
        }
    }
}
