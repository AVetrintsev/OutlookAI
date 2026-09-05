using OutlookAI.TaskPane.InboxCopilot;
using Xunit;

namespace OutlookAI.Tests.TaskPane.InboxCopilot
{
    public class QuickActionChipTests
    {
        [Fact]
        public void NoSelection_ReturnsThreeStaticChips()
        {
            var chips = QuickActionChip.ComputeChipsForSelectionCount(0);
            Assert.Equal(3, chips.Count);
            Assert.Contains(chips, c => c.Label == "Что требует внимания?");
            Assert.Contains(chips, c => c.Label == "Сводка непрочитанных");
            Assert.Contains(chips, c => c.Label == "Письма за сегодня");
        }

        [Fact]
        public void SingleSelection_AddsSingleSelectionChips()
        {
            var chips = QuickActionChip.ComputeChipsForSelectionCount(1);
            Assert.Equal(5, chips.Count);
            Assert.Contains(chips, c => c.Label == "Сводка переписки");
            Assert.Contains(chips, c => c.Label == "Черновик ответа");
            // Static three still present:
            Assert.Contains(chips, c => c.Label == "Что требует внимания?");
            Assert.Contains(chips, c => c.Label == "Сводка непрочитанных");
            Assert.Contains(chips, c => c.Label == "Письма за сегодня");
            Assert.DoesNotContain(chips, c => c.Label == "Сводка выбранных");
            Assert.DoesNotContain(chips, c => c.Label == "Разобрать выбранные");
        }

        [Fact]
        public void MultiSelection_AddsMultiSelectionChips()
        {
            var chips = QuickActionChip.ComputeChipsForSelectionCount(3);
            Assert.Equal(5, chips.Count);
            Assert.Contains(chips, c => c.Label == "Сводка выбранных");
            Assert.Contains(chips, c => c.Label == "Разобрать выбранные");
            Assert.Contains(chips, c => c.Label == "Что требует внимания?");
            Assert.Contains(chips, c => c.Label == "Сводка непрочитанных");
            Assert.Contains(chips, c => c.Label == "Письма за сегодня");
            Assert.DoesNotContain(chips, c => c.Label == "Сводка переписки");
            Assert.DoesNotContain(chips, c => c.Label == "Черновик ответа");
        }

        [Fact]
        public void Prompts_AreNotEmpty()
        {
            foreach (var n in new[] { 0, 1, 5 })
            {
                var chips = QuickActionChip.ComputeChipsForSelectionCount(n);
                Assert.All(chips, c => Assert.False(string.IsNullOrWhiteSpace(c.Prompt)));
                Assert.All(chips, c => Assert.False(string.IsNullOrWhiteSpace(c.Label)));
            }
        }
    }
}
