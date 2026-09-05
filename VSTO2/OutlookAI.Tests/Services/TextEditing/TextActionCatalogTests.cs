using System.Linq;
using OutlookAI.Services;
using OutlookAI.Services.CustomActions;
using OutlookAI.Services.TextEditing;
using Xunit;

namespace OutlookAI.Tests.Services.TextEditing
{
    public sealed class TextActionCatalogTests
    {
        [Fact]
        public void IsTextAction_RequiresSurfaceSourceAndOutputContract()
        {
            var valid = Action("valid");
            Assert.True(TextActionCatalog.IsTextAction(valid));

            var assistant = Action("assistant");
            assistant.Surface = CustomActionSurface.Assistant;
            Assert.False(TextActionCatalog.IsTextAction(assistant));

            var wrongSource = Action("source");
            wrongSource.Context.Source = "current_selection";
            Assert.False(TextActionCatalog.IsTextAction(wrongSource));

            var wrongOutput = Action("output");
            wrongOutput.Output = "chat";
            Assert.False(TextActionCatalog.IsTextAction(wrongOutput));
        }

        [Fact]
        public void IsTextAction_NormalizesContractCasingAndWhitespace()
        {
            var action = Action("normalized");
            action.Surface = " EDITOR_SELECTION ";
            action.Context.Source = " SELECTED_TEXT ";
            action.Output = " REPLACE_SELECTION ";

            Assert.True(TextActionCatalog.IsTextAction(action));
        }

        [Fact]
        public void GetActions_FiltersAndPreservesGroupThenActionOrder()
        {
            var first = Action("first");
            var second = Action("second");
            var last = Action("last");
            var disabled = Action("disabled");
            disabled.Disabled = true;
            var meetingOnly = Action("meeting-only");
            meetingOnly.ApplicabilityItemType = CustomActionApplicability.Meeting;
            var assistant = Action("assistant");
            assistant.Surface = CustomActionSurface.Assistant;

            var catalog = new CustomActionFile
            {
                Groups = new[]
                {
                    new CustomActionGroup
                    {
                        Id = "later",
                        Title = "Later",
                        Order = 20,
                        Actions = new[] { last }
                    },
                    new CustomActionGroup
                    {
                        Id = "earlier",
                        Title = "Earlier",
                        Order = 10,
                        Actions = new[] { first, disabled, second, meetingOnly, assistant }
                    }
                }
            };

            var actions = TextActionCatalog.GetActions(catalog, "mail", "outgoing");

            Assert.Equal(new[] { "first", "second", "last" }, actions.Select(item => item.Id));
        }

        [Fact]
        public void GetActions_AppliesDirectionAndSupportsAllForTasks()
        {
            var all = Action("all");
            var incoming = Action("incoming");
            incoming.ApplicabilityDirection = CustomActionApplicability.Incoming;
            var outgoing = Action("outgoing");
            outgoing.ApplicabilityDirection = CustomActionApplicability.Outgoing;
            var catalog = FileWith(all, incoming, outgoing);

            var taskActions = TextActionCatalog.GetActions(catalog, "task", "outgoing");

            Assert.Equal(new[] { "all", "outgoing" }, taskActions.Select(item => item.Id));
        }

        [Fact]
        public void GetActions_ReturnsDefensiveClones()
        {
            var original = Action("one");
            var catalog = FileWith(original);

            var returned = Assert.Single(TextActionCatalog.GetActions(catalog, "mail", "incoming"));
            returned.Title = "Changed";
            returned.Context.Source = "Changed";

            Assert.Equal("one", original.Title);
            Assert.Equal(TextActionCatalog.Source, original.Context.Source);
        }

        [Fact]
        public void GetActions_NullCatalog_ReturnsEmpty()
        {
            Assert.Empty(TextActionCatalog.GetActions(null, "mail", "incoming"));
        }

        [Fact]
        public void BuiltinCatalog_ProvidesTheFiveDefaultEditorActionsInOrder()
        {
            var catalog = new CustomActionFile
            {
                Groups = ActionCatalog.Default.AssistantGroups().ToArray()
            };

            var actions = TextActionCatalog.GetActions(catalog, "task", "outgoing");

            Assert.Equal(
                new[]
                {
                    "text_editing_01",
                    "text_editing_02",
                    "text_editing_03",
                    "text_editing_04",
                    "text_editing_06",
                    "text_editing_07",
                    "text_editing_05"
                },
                actions.Select(action => action.Id));
        }

        private static CustomActionFile FileWith(params CustomActionDefinition[] actions)
        {
            return new CustomActionFile
            {
                Groups = new[]
                {
                    new CustomActionGroup
                    {
                        Id = "text",
                        Title = "Text",
                        Actions = actions
                    }
                }
            };
        }

        private static CustomActionDefinition Action(string id)
        {
            return new CustomActionDefinition
            {
                Id = id,
                Title = id,
                Prompt = "Rewrite clearly.",
                Surface = TextActionCatalog.Surface,
                Context = new CustomActionContext { Source = TextActionCatalog.Source },
                Output = TextActionCatalog.Output,
                ApplicabilityItemType = CustomActionApplicability.All,
                ApplicabilityDirection = CustomActionApplicability.All,
                AllowedTools = new string[0]
            };
        }
    }
}
