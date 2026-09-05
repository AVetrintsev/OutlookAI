using System;
using System.Linq;
using OutlookAI.Services.TextEditing;
using Xunit;

namespace OutlookAI.Tests.Services.TextEditing
{
    public sealed class ProtectedTextPlanTests
    {
        [Fact]
        public void LinksRemainExactlyOnceInOrder()
        {
            var plan = new ProtectedTextPlan(new[] { "Прочти ", " и ", "." }, new[] { "документ", "план" });
            var edited = plan.ModelText.Replace("Прочти", "Пожалуйста, изучите");
            Assert.Equal("Пожалуйста, изучите документ и план.", plan.GetText(plan.ParseReplacement(edited)));
            Assert.Throws<InvalidOperationException>(() => plan.ParseReplacement("Ссылок нет"));
            Assert.Throws<InvalidOperationException>(() => plan.ParseReplacement(plan.ModelText + plan.ModelText));
            var tokens = System.Text.RegularExpressions.Regex.Matches(plan.ModelText, @"\[\[OUTLOOKAI_LINK_[^\]]+\]\]");
            Assert.Throws<InvalidOperationException>(() => plan.ParseReplacement(tokens[1].Value + tokens[0].Value));
        }

        [Theory]
        [InlineData("Текст без изменений", "Текст без изменений")]
        [InlineData("Привет, мир!", "Здравствуйте, мир!")]
        [InlineData("a\rb\rc", "a\rc")]
        [InlineData("", "новый текст")]
        [InlineData("старый текст", "")]
        [InlineData("😀 цветной текст", "🟢 цветной текст")]
        [InlineData("одно одно два", "одно два одно")]
        public void PatchesReconstructReplacement(string original, string expected)
        {
            var actual = original;
            foreach (var change in TextChange.Calculate(original, expected).Reverse())
                actual = actual.Remove(change.Start, change.Length).Insert(change.Start, change.Text);
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void UnchangedWordsAreNotReplaced()
        {
            var changes = TextChange.Calculate("Сохранить жирный текст и ссылку", "Сохранить жирный текст и гиперссылку");
            var change = Assert.Single(changes);
            Assert.Equal("ссылку", "Сохранить жирный текст и ссылку".Substring(change.Start, change.Length));
            Assert.Equal("гиперссылку", change.Text);
        }

        [Fact]
        public void RandomEditsReconstructExactly()
        {
            var random = new Random(42);
            for (var i = 0; i < 300; i++)
            {
                Func<string> make = () => string.Join(" ", Enumerable.Range(0, random.Next(1, 35)).Select(_ => "w" + random.Next(6)));
                PatchesReconstructReplacement(make(), make());
            }
        }

        [Fact]
        public void OversizedDiffIsRefusedBeforeAnyWrite()
        {
            Assert.Throws<InvalidOperationException>(() => TextChange.Calculate(
                string.Join(" ", Enumerable.Repeat("old", 2000)), string.Join(" ", Enumerable.Repeat("new", 2000))));
        }
    }
}
