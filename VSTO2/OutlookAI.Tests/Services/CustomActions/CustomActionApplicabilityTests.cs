using System.Linq;
using OutlookAI.Services;
using OutlookAI.Services.CustomActions;
using Xunit;

namespace OutlookAI.Tests.Services.CustomActions
{
    public sealed class CustomActionApplicabilityTests
    {
        [Fact]
        public void IsApplicable_MailIncomingAction_OnlyMatchesIncomingMail()
        {
            var action = Action("mail", "incoming");

            Assert.True(CustomActionApplicability.IsApplicable(action, Context("mail", "incoming")));
            Assert.False(CustomActionApplicability.IsApplicable(action, Context("mail", "outgoing")));
            Assert.False(CustomActionApplicability.IsApplicable(action, Context("meeting", "incoming")));
        }

        [Fact]
        public void IsApplicable_MeetingOutgoingAction_OnlyMatchesOutgoingMeeting()
        {
            var action = Action("meeting", "outgoing");

            Assert.True(CustomActionApplicability.IsApplicable(action, Context("meeting", "outgoing")));
            Assert.False(CustomActionApplicability.IsApplicable(action, Context("meeting", "incoming")));
            Assert.False(CustomActionApplicability.IsApplicable(action, Context("mail", "outgoing")));
        }

        [Fact]
        public void FilterCatalog_RemovesEmptyGroups()
        {
            var catalog = new CustomActionFile
            {
                Groups = new[]
                {
                    new CustomActionGroup
                    {
                        Id = "g1",
                        Actions = new[] { Action("mail", "incoming") }
                    },
                    new CustomActionGroup
                    {
                        Id = "g2",
                        Actions = new[] { Action("meeting", "outgoing") }
                    }
                }
            };

            var filtered = CustomActionApplicability.FilterCatalog(catalog, Context("mail", "incoming"));

            Assert.Single(filtered.Groups);
            Assert.Equal("g1", filtered.Groups[0].Id);
        }

        [Fact]
        public void BuiltinGroupedActions_AllHaveValidApplicability()
        {
            var groups = ActionCatalog.Default.AssistantGroups();
            Assert.NotEmpty(groups);
            foreach (var action in groups.SelectMany(group => group.Actions))
            {
                Assert.Contains(
                    CustomActionApplicability.NormalizeItemType(action.ApplicabilityItemType),
                    new[] { "all", "mail", "meeting" });
                Assert.Contains(
                    CustomActionApplicability.NormalizeDirection(action.ApplicabilityDirection),
                    new[] { "all", "incoming", "outgoing" });
            }
        }

        [Theory]
        [InlineData("reply_01", "mail", "incoming")]
        [InlineData("reply_05", "mail", "outgoing")]
        [InlineData("meetings_03", "meeting", "all")]
        [InlineData("control_03", "mail", "outgoing")]
        [InlineData("understand_01", "all", "all")]
        public void BuiltinGroupedActions_HaveExpectedApplicability(string id, string itemType, string direction)
        {
            var action = ActionCatalog.Default.AssistantGroups()
                .SelectMany(group => group.Actions)
                .Single(candidate => candidate.Id == id);

            Assert.Equal(itemType, action.ApplicabilityItemType);
            Assert.Equal(direction, action.ApplicabilityDirection);
        }

        private static CustomActionDefinition Action(string itemType, string direction)
        {
            return new CustomActionDefinition
            {
                Id = itemType + "_" + direction,
                Title = "Action",
                Prompt = "Prompt",
                ApplicabilityItemType = itemType,
                ApplicabilityDirection = direction
            };
        }

        private static CustomActionApplicabilityContext Context(string itemType, string direction)
        {
            return new CustomActionApplicabilityContext
            {
                ItemType = itemType,
                Direction = direction
            };
        }
    }
}
