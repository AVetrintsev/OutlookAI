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
                Assert.Equal(8, catalog.Groups.Count(group => group.Id != "my_actions"));
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

                Assert.Equal(8, catalog.Groups.Length);
                Assert.All(catalog.Groups, group => Assert.NotEmpty(group.Actions));
                Assert.Equal(8,
                    JsonConvert.DeserializeObject<CustomActionFile>(File.ReadAllText(path)).Groups.Length);
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
