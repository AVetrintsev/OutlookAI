using System.Linq;
using Newtonsoft.Json.Linq;
using OutlookAI.Services.CustomActions;
using Xunit;

namespace OutlookAI.Tests.Services.CustomActions
{
    public sealed class CustomActionUiSerializerTests
    {
        [Fact]
        public void Build_CanExcludeEditorSelectionActionsAndRecommendations()
        {
            var catalog = new CustomActionFile
            {
                Groups = new[]
                {
                    Group("inline_group", Action("inline", CustomActionSurface.EditorSelection)),
                    Group("assistant_group", Action("assistant", CustomActionSurface.Assistant))
                }
            };

            var json = CustomActionUiSerializer.Build(
                catalog,
                new[] { "inline", "assistant" },
                new JArray(),
                new CustomActionApplicabilityContext
                {
                    ItemType = CustomActionApplicability.Mail,
                    Direction = CustomActionApplicability.Incoming
                },
                includeEditorSelectionActions: false);

            var group = Assert.Single(((JArray)json["groups"]).OfType<JObject>());
            Assert.Equal("assistant_group", (string)group["id"]);
            Assert.Equal("assistant", Assert.Single(
                ((JArray)json["recommendation_ids"]).Values<string>()));
        }

        [Fact]
        public void Build_PreservesEmptyTextEditingGroupOnlyForComposeSurface()
        {
            var catalog = new CustomActionFile
            {
                Groups = new[]
                {
                    new CustomActionGroup
                    {
                        Id = CustomActionStore.TextEditingGroupId,
                        Title = "Редактирование текста",
                        Order = 0,
                        Actions = new CustomActionDefinition[0]
                    },
                    new CustomActionGroup
                    {
                        Id = "other_empty",
                        Title = "Другая пустая группа",
                        Order = 1,
                        Actions = new CustomActionDefinition[0]
                    }
                }
            };

            var composeJson = CustomActionUiSerializer.Build(
                catalog,
                new string[0],
                new JArray());
            var composeGroup = Assert.Single(((JArray)composeJson["groups"]).OfType<JObject>());
            Assert.Equal(CustomActionStore.TextEditingGroupId, (string)composeGroup["id"]);
            Assert.Equal("Редактирование текста", (string)composeGroup["title"]);
            Assert.True((bool)composeGroup["can_reset"]);
            Assert.Empty((JArray)composeGroup["actions"]);

            var explorerJson = CustomActionUiSerializer.Build(
                catalog,
                new string[0],
                new JArray(),
                includeEditorSelectionActions: false);
            Assert.Empty((JArray)explorerJson["groups"]);
        }

        [Fact]
        public void ToJson_SerializesSurfaceAndTaskApplicability()
        {
            var action = Action("inline", CustomActionSurface.EditorSelection);
            action.ApplicabilityItemType = CustomActionApplicability.Task;

            var json = CustomActionUiSerializer.ToJson(action);

            Assert.Equal(CustomActionSurface.EditorSelection, (string)json["surface"]);
            Assert.Equal(CustomActionApplicability.Task, (string)json["applicability_item_type"]);
        }

        private static CustomActionGroup Group(string id, CustomActionDefinition action)
        {
            return new CustomActionGroup
            {
                Id = id,
                Title = id,
                Actions = new[] { action }
            };
        }

        private static CustomActionDefinition Action(string id, string surface)
        {
            return new CustomActionDefinition
            {
                Id = id,
                Title = id,
                Prompt = id,
                Surface = surface,
                Context = new CustomActionContext { Source = "current_selection" },
                Output = "chat",
                ApplicabilityItemType = CustomActionApplicability.All,
                ApplicabilityDirection = CustomActionApplicability.All
            };
        }
    }
}
