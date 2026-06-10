using System.Collections.Generic;

namespace OutlookAI.TaskPane.InboxReports
{
    /// <summary>
    /// One quick-action chip in the Inbox Reports pane. Clicking the chip
    /// prefills the chat input with TemplateText. The user can edit
    /// any [placeholders] before sending; the system prompt instructs
    /// the model to ask for clarification if a placeholder is still
    /// present at tool-call time.
    /// </summary>
    public sealed class ReportQuickActionChip
    {
        public string Label { get; set; }
        public string TemplateText { get; set; }

        public static IReadOnlyList<ReportQuickActionChip> Defaults()
        {
            return new[]
            {
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCC5 Дайджест недели",
                    TemplateText = "Сделай сводку того, что пришло во Входящие за последние 7 дней. Сгруппируй по отправителю или теме. Выдели срочные письма и письма, адресованные лично мне.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCAC Сводка переписки",
                    TemplateText = "Сделай сводку моих последних переписок с [имя или email]. Покажи хронологию и ключевые решения/темы.",
                },
                new ReportQuickActionChip {
                    Label = "\u2713 Задачи",
                    TemplateText = "Найди задачи, которые мне нужно выполнить по письмам за последние 7 дней. Прочитай релевантные сообщения, извлеки TODO/сроки/запросы. Сгруппируй по тому, кто чего ждёт.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCC1 Статус проекта",
                    TemplateText = "Сделай сводку статуса по [тема/название проекта]. Найди релевантные письма, прочитай самые свежие и дай: последнее обновление, открытые вопросы, задачи, ключевых участников.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83D\uDCCA Статистика почты",
                    TemplateText = "Дай статистику почты за последние 30 дней: топ-10 отправителей, самые загруженные дни, разбивку по папкам. Используй outlook_aggregate_messages.",
                },
                new ReportQuickActionChip {
                    Label = "\uD83C\uDFD6\uFE0F Пока меня не было",
                    TemplateText = "Меня не было с [дата начала] по [дата окончания]. Покажи важное за этот период: срочные вопросы, прямые запросы, письма, на которые нужен ответ. Рассылки и автоматические письма поставь ниже по приоритету.",
                },
            };
        }
    }
}
