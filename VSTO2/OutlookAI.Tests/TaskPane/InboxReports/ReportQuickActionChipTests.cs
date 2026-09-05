using System.Linq;
using OutlookAI.TaskPane.InboxReports;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxReports
{
    public class ReportQuickActionChipTests
    {
        private readonly System.Collections.Generic.IReadOnlyList<ReportQuickActionChip> _chips
            = ReportQuickActionChip.Defaults();

        [Fact]
        public void Defaults_ReturnsSixChips()
        {
            Assert.Equal(6, _chips.Count);
        }

        [Fact]
        public void Defaults_EachChipHasLabelAndTemplate()
        {
            foreach (var c in _chips)
            {
                Assert.False(string.IsNullOrWhiteSpace(c.Label));
                Assert.False(string.IsNullOrWhiteSpace(c.TemplateText));
                Assert.True(c.TemplateText.Length > 40,
                    "Chip template too short: " + c.Label);
            }
        }

        [Fact]
        public void Defaults_OrderingMatchesSpec()
        {
            // Spec defines the order: Digest, Conversation, Action items,
            // Project status, Stats, Out-of-office.
            Assert.Equal(new[]
            {
                "📅 Дайджест недели",
                "💬 Сводка переписки",
                "✓ Задачи",
                "📁 Статус проекта",
                "📊 Статистика почты",
                "🏖️ Пока меня не было"
            }, _chips.Select(c => c.Label));
        }

        [Fact]
        public void ConversationChip_TemplateMentionsPersonPlaceholder()
        {
            Assert.Contains("[имя или email]", _chips[1].TemplateText);
        }

        [Fact]
        public void ProjectChip_TemplateMentionsTopicPlaceholder()
        {
            Assert.Contains("[тема/название проекта]", _chips[3].TemplateText);
        }

        [Fact]
        public void OutOfOfficeChip_TemplateMentionsDatePlaceholders()
        {
            Assert.Contains("[дата начала]", _chips[5].TemplateText);
            Assert.Contains("[дата окончания]", _chips[5].TemplateText);
        }

        [Fact]
        public void StatsChip_TemplateMentionsAggregateTool()
        {
            Assert.Contains("outlook_aggregate_messages", _chips[4].TemplateText);
        }

        [Fact]
        public void Defaults_LabelsAreUnique()
        {
            Assert.Equal(_chips.Count, _chips.Select(c => c.Label).Distinct().Count());
        }
    }
}
