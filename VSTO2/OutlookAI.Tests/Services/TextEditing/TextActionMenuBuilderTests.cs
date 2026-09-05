using System.Linq;
using System.Xml.Linq;
using OutlookAI.Services.CustomActions;
using OutlookAI.Services.TextEditing;
using Xunit;

namespace OutlookAI.Tests.Services.TextEditing
{
    public sealed class TextActionMenuBuilderTests
    {
        private static readonly XNamespace RibbonNs = TextActionMenuBuilder.RibbonNamespace;

        [Fact]
        public void Build_EmptyActions_ReturnsValidEmptyMenu()
        {
            var document = XDocument.Parse(TextActionMenuBuilder.Build(null));

            Assert.Equal(RibbonNs + "menu", document.Root.Name);
            Assert.Empty(document.Root.Elements(RibbonNs + "button"));
        }

        [Fact]
        public void Build_UsesEscapedAttributesTagAndAgreedCallback()
        {
            var action = Action("id&<\"quoted", "A & <B> \"quote\" \u0001 😀");
            action.Description = "Hint & <safe>";

            var xml = TextActionMenuBuilder.Build(new[] { action });
            var button = Assert.Single(XDocument.Parse(xml).Root.Elements(RibbonNs + "button"));

            Assert.Equal("id&<\"quoted", (string)button.Attribute("tag"));
            Assert.Equal("A & <B> \"quote\"  😀", (string)button.Attribute("label"));
            Assert.Equal("Hint & <safe>", (string)button.Attribute("supertip"));
            Assert.Equal(TextActionMenuBuilder.OnActionCallback, (string)button.Attribute("onAction"));
            Assert.Contains("&amp;", xml);
            Assert.Contains("&lt;", xml);
            Assert.DoesNotContain("\u0001", xml);
        }

        [Fact]
        public void Build_PreservesActionOrderAndGeneratesSafeUniqueControlIds()
        {
            var document = XDocument.Parse(TextActionMenuBuilder.Build(new[]
            {
                Action("z", "Last id first"),
                Action("a", "First id second")
            }));
            var buttons = document.Root.Elements(RibbonNs + "button").ToArray();

            Assert.Equal(new[] { "z", "a" }, buttons.Select(button => (string)button.Attribute("tag")));
            Assert.Equal(
                new[] { "outlookAiTextAction0", "outlookAiTextAction1" },
                buttons.Select(button => (string)button.Attribute("id")));
        }

        [Fact]
        public void Build_FiltersDisabledNonTextDuplicateAndInvalidIdActions()
        {
            var valid = Action("valid", "Valid");
            var duplicate = Action("VALID", "Duplicate");
            var disabled = Action("disabled", "Disabled");
            disabled.Disabled = true;
            var assistant = Action("assistant", "Assistant");
            assistant.Surface = CustomActionSurface.Assistant;
            var invalidId = Action("bad\u0001id", "Invalid id");

            var document = XDocument.Parse(TextActionMenuBuilder.Build(new[]
            {
                valid, duplicate, disabled, assistant, invalidId
            }));
            var button = Assert.Single(document.Root.Elements(RibbonNs + "button"));

            Assert.Equal("valid", (string)button.Attribute("tag"));
        }

        [Fact]
        public void Build_LimitsVeryLongLabelsWithoutSplittingEmoji()
        {
            var action = Action("long", new string('x', 254) + "😀" + "tail");

            var button = Assert.Single(XDocument.Parse(
                TextActionMenuBuilder.Build(new[] { action }))
                .Root.Elements(RibbonNs + "button"));
            var label = (string)button.Attribute("label");

            Assert.Equal(254, label.Length);
            Assert.DoesNotContain("tail", label);
        }

        private static CustomActionDefinition Action(string id, string title)
        {
            return new CustomActionDefinition
            {
                Id = id,
                Title = title,
                Description = "",
                Prompt = "Rewrite",
                Surface = TextActionCatalog.Surface,
                Context = new CustomActionContext { Source = TextActionCatalog.Source },
                Output = TextActionCatalog.Output,
                AllowedTools = new string[0]
            };
        }
    }
}
