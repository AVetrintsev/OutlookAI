using System.Linq;
using OutlookAI.Services.CustomActions;
using Xunit;

namespace OutlookAI.Tests.Services.CustomActions
{
    public sealed class CustomActionYamlCodecTests
    {
        [Fact]
        public void ExportParse_RoundTripsContextAndTools()
        {
            var file = FileWith("Понять", "Суть письма", "old");
            file.Groups[0].Actions[0].AllowedTools = new[] { "outlook_read_message" };
            file.Groups[0].Actions[0].AllowTools = true;

            var yaml = CustomActionYamlCodec.Export(file);
            var parsed = CustomActionYamlCodec.Parse(yaml);
            var action = Assert.Single(Assert.Single(parsed.Groups).Actions);

            Assert.Equal("Суть письма", action.Title);
            Assert.Equal("related_thread", action.Context.Source);
            Assert.Equal("mail", action.ApplicabilityItemType);
            Assert.Equal("incoming", action.ApplicabilityDirection);
            Assert.Equal("outlook_read_message", Assert.Single(action.AllowedTools));
        }

        [Fact]
        public void PreviewAndMerge_MatchByGroupAndActionTitle()
        {
            var current = FileWith("Понять", "Суть письма", "old");
            var imported = FileWith("Понять", "Суть письма", "new");
            imported.Groups[0].Actions[0].Id = "different-id";
            var yaml = CustomActionYamlCodec.Export(imported);

            var preview = CustomActionYamlCodec.Preview(yaml, current);
            var merged = CustomActionYamlCodec.Merge(current, preview.Imported);
            var action = Assert.Single(Assert.Single(merged.Groups).Actions);

            Assert.Equal("Понять / Суть письма", Assert.Single(preview.Replacements));
            Assert.Equal("old", action.Id);
            Assert.Equal("new", action.Prompt);
        }

        [Fact]
        public void Parse_AcceptsCompactGroupItemsFormat()
        {
            var parsed = CustomActionYamlCodec.Parse(
                "- title: Понять\n" +
                "  items:\n" +
                "    - title: Суть письма\n" +
                "      description: Краткое описание\n" +
                "      prompt: >\n" +
                "        Разбери письмо.\n");

            var action = Assert.Single(Assert.Single(parsed.Groups).Actions);
            Assert.Equal("Понять", parsed.Groups[0].Title);
            Assert.Equal("Суть письма", action.Title);
            Assert.Equal("Разбери письмо.", action.Prompt);
        }

        private static CustomActionFile FileWith(string groupTitle, string actionTitle, string prompt)
        {
            return new CustomActionFile
            {
                SchemaVersion = 2,
                Groups = new[]
                {
                    new CustomActionGroup
                    {
                        Id = "group",
                        Title = groupTitle,
                        Actions = new[]
                        {
                            new CustomActionDefinition
                            {
                                Id = "old",
                                Title = actionTitle,
                                Description = "Описание",
                                Prompt = prompt,
                                Context = new CustomActionContext
                                {
                                    Source = "related_thread",
                                    IncludeFullBodies = true,
                                    MaxItems = 20
                                },
                                Output = "chat",
                                ApplicabilityItemType = "mail",
                                ApplicabilityDirection = "incoming"
                            }
                        }
                    }
                }
            };
        }
    }
}
