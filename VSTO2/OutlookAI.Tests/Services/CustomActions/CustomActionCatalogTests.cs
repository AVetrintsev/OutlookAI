using System.Linq;
using OutlookAI.Services;
using OutlookAI.Services.CustomActions;
using Xunit;

namespace OutlookAI.Tests.Services.CustomActions
{
    public sealed class CustomActionCatalogTests
    {
        [Fact]
        public void BuiltinCatalog_TextEditingGroup_HasSevenEditorSelectionActions()
        {
            var groups = ActionCatalog.Default.AssistantGroups();
            var group = Assert.Single(
                groups,
                item => item.Id == CustomActionStore.TextEditingGroupId);

            Assert.Equal(0, group.Order);
            Assert.Equal(7, group.Actions.Length);
            Assert.Equal(
                new[]
                {
                    "Официально",
                    "Структурировать",
                    "Уточнить",
                    "Сослаться на пункты переписки",
                    "Строго и прямо",
                    "Мягко и дипломатично",
                    "Кратко"
                },
                group.Actions.Select(action => action.Title));
            Assert.All(group.Actions, action =>
            {
                Assert.Equal(CustomActionSurface.EditorSelection, action.Surface);
                Assert.Equal("selected_text", action.Context.Source);
                Assert.Equal("selected", action.Context.MessageScope);
                Assert.False(action.Context.IncludeFullBodies);
                Assert.False(action.Context.IncludeAttachments);
                Assert.Equal(1, action.Context.MaxItems);
                Assert.Equal("replace_selection", action.Output);
                Assert.False(action.AllowTools);
                Assert.Empty(action.AllowedTools);
            });
        }

        [Fact]
        public void BuiltinCatalog_ParsesSnakeCaseContextFields()
        {
            var action = ActionCatalog.Default.AssistantGroups()
                .SelectMany(group => group.Actions)
                .Single(item => item.Id == "understand_01");

            Assert.Equal("related_thread", action.Context.Source);
            Assert.Equal("thread", action.Context.MessageScope);
            Assert.Equal("current_folder", action.Context.FolderScope);
            Assert.Equal("all", action.Context.ReadFilter);
            Assert.Equal("today", action.Context.TimeRange);
            Assert.True(action.Context.IncludeFullBodies);
            Assert.False(action.Context.IncludeAttachments);
            Assert.Equal(100, action.Context.MaxItems);
        }
    }
}
