using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using OutlookAI.Services.CustomActions;
using Xunit;

namespace OutlookAI.Tests.Services.CustomActions
{
    public sealed class CustomActionStoreTests
    {
        [Fact]
        public void LoadCatalog_MigratesFlatSchemaIntoPersonalGroup()
        {
            var path = TempFile();
            try
            {
                File.WriteAllText(path, JsonConvert.SerializeObject(new CustomActionFile
                {
                    Actions = new[] { Action("legacy", "Старое действие") }
                }));

                var catalog = new CustomActionStore(path).LoadCatalog();

                Assert.Equal(CustomActionStore.CurrentSchemaVersion, catalog.SchemaVersion);
                Assert.Equal(
                    OutlookAI.Services.ActionCatalog.Default.AssistantGroups().Count,
                    catalog.Groups.Count(group => group.Id != "my_actions"));
                var personal = Assert.Single(catalog.Groups, group => group.Id == "my_actions");
                Assert.Equal("legacy", Assert.Single(personal.Actions).Id);
                Assert.Null(JsonConvert.DeserializeObject<CustomActionFile>(File.ReadAllText(path)).Actions);
            }
            finally
            {
                TryDelete(path);
            }
        }

        [Fact]
        public void ReorderAndResetGroup_PersistWorkingCatalog()
        {
            var path = TempFile();
            try
            {
                var store = new CustomActionStore(path);
                var catalog = store.LoadCatalog();
                var group = catalog.Groups.First();
                var reversed = group.Actions.Select(action => action.Id).Reverse().ToArray();

                Assert.True(store.Reorder(group.Id, reversed));
                Assert.Equal(reversed, store.LoadCatalog().Groups.First().Actions.Select(action => action.Id));

                Assert.True(store.ResetGroup(group.Id));
                Assert.Equal(
                    OutlookAI.Services.ActionCatalog.Default.AssistantGroups().First().Actions.Select(action => action.Id),
                    store.LoadCatalog().Groups.First().Actions.Select(action => action.Id));
            }
            finally
            {
                TryDelete(path);
            }
        }

        [Fact]
        public void LoadCatalog_RepairsEmptyGroupedSchema()
        {
            var path = TempFile();
            try
            {
                File.WriteAllText(path, JsonConvert.SerializeObject(new CustomActionFile
                {
                    SchemaVersion = CustomActionStore.CurrentSchemaVersion,
                    Groups = new CustomActionGroup[0]
                }));

                var catalog = new CustomActionStore(path).LoadCatalog();

                var expectedGroupCount = OutlookAI.Services.ActionCatalog.Default.AssistantGroups().Count;
                Assert.Equal(expectedGroupCount, catalog.Groups.Length);
                Assert.All(catalog.Groups, group => Assert.NotEmpty(group.Actions));
                Assert.Equal(expectedGroupCount,
                    JsonConvert.DeserializeObject<CustomActionFile>(File.ReadAllText(path)).Groups.Length);
            }
            finally
            {
                TryDelete(path);
            }
        }

        [Fact]
        public void LoadCatalog_InsertsNewTextEditingGroupAtBaselinePosition()
        {
            var path = TempFile();
            try
            {
                var oldGroups = OutlookAI.Services.ActionCatalog.Default.AssistantGroups()
                    .Where(group => group.Id != CustomActionStore.TextEditingGroupId)
                    .Select((group, index) =>
                    {
                        group.Order = index;
                        return group;
                    })
                    .ToArray();
                File.WriteAllText(path, JsonConvert.SerializeObject(new CustomActionFile
                {
                    SchemaVersion = CustomActionStore.CurrentSchemaVersion,
                    Groups = oldGroups
                }));

                var catalog = new CustomActionStore(path).LoadCatalog();

                Assert.Equal(CustomActionStore.TextEditingGroupId, catalog.Groups[0].Id);
                Assert.Equal(
                    OutlookAI.Services.ActionCatalog.Default.AssistantGroups().Select(group => group.Id),
                    catalog.Groups.Select(group => group.Id));
            }
            finally
            {
                TryDelete(path);
            }
        }

        [Fact]
        public void Upsert_TextEditingAction_EnforcesEditorSelectionContract()
        {
            var path = TempFile();
            try
            {
                var store = new CustomActionStore(path);
                var action = Action("custom_inline", "Моё редактирование");

                store.Upsert(CustomActionStore.TextEditingGroupId, action);

                var saved = store.LoadCatalog().Groups
                    .Single(group => group.Id == CustomActionStore.TextEditingGroupId)
                    .Actions.Single(item => item.Id == action.Id);
                Assert.Equal(CustomActionSurface.EditorSelection, saved.Surface);
                Assert.Equal("selected_text", saved.Context.Source);
                Assert.Equal("replace_selection", saved.Output);
                Assert.Equal(1, saved.Context.MaxItems);
                Assert.False(saved.Context.IncludeFullBodies);
                Assert.False(saved.Context.IncludeAttachments);
                Assert.False(saved.AllowTools);
                Assert.Empty(saved.AllowedTools);
            }
            finally
            {
                TryDelete(path);
            }
        }

        [Fact]
        public void LoadCatalog_PreservesPascalCaseContextFromExistingUserFile()
        {
            var path = TempFile();
            try
            {
                File.WriteAllText(path, JsonConvert.SerializeObject(new CustomActionFile
                {
                    SchemaVersion = CustomActionStore.CurrentSchemaVersion,
                    Groups = new[]
                    {
                        new CustomActionGroup
                        {
                            Id = "existing_custom",
                            Title = "Пользовательская группа",
                            Actions = new[]
                            {
                                new CustomActionDefinition
                                {
                                    Id = "existing_action",
                                    Title = "Действие",
                                    Prompt = "Промпт",
                                    Surface = CustomActionSurface.Assistant,
                                    Context = new CustomActionContext
                                    {
                                        Source = "all_folders",
                                        MessageScope = "selected",
                                        FolderScope = "all_folders",
                                        ReadFilter = "unread",
                                        TimeRange = "yesterday",
                                        IncludeFullBodies = false,
                                        IncludeAttachments = true,
                                        MaxItems = 37
                                    },
                                    Output = "chat"
                                }
                            }
                        }
                    }
                }));

                var action = new CustomActionStore(path).LoadCatalog().Groups
                    .Single(group => group.Id == "existing_custom")
                    .Actions.Single();

                Assert.Equal("all_folders", action.Context.Source);
                Assert.Equal("selected", action.Context.MessageScope);
                Assert.Equal("all_folders", action.Context.FolderScope);
                Assert.Equal("unread", action.Context.ReadFilter);
                Assert.Equal("yesterday", action.Context.TimeRange);
                Assert.False(action.Context.IncludeFullBodies);
                Assert.True(action.Context.IncludeAttachments);
                Assert.Equal(37, action.Context.MaxItems);
            }
            finally
            {
                TryDelete(path);
            }
        }

        private static CustomActionDefinition Action(string id, string title)
        {
            return new CustomActionDefinition
            {
                Id = id,
                Title = title,
                Prompt = "Промпт",
                Context = new CustomActionContext
                {
                    Source = "current_selection",
                    IncludeFullBodies = true,
                    MaxItems = 20
                },
                Output = "chat"
            };
        }

        private static string TempFile()
        {
            var path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "outlookai-actions-" + Guid.NewGuid().ToString("N") + ".json");
            return path;
        }

        private static void TryDelete(string path)
        {
            try { File.Delete(path); } catch { }
        }
    }
}
