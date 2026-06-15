using System.IO;
using Xunit;

namespace OutlookAI.Tests.TaskPane.Chat
{
    public sealed class ActionRecommendationWebUiTests
    {
        [Fact]
        public void GroupButtonsHaveNoDecorativeArrow()
        {
            var css = ReadWebUiFile("styles.css");

            Assert.DoesNotContain(".action-group-btn::after", css);
            Assert.DoesNotContain("padding-right: 24px", css);
        }

        [Fact]
        public void RecommendationUiSupportsExplicitStates()
        {
            var script = ReadWebUiFile("chat.js");

            Assert.Contains("setActionRecommendationState", script);
            Assert.Contains("state === 'disabled'", script);
            Assert.Contains("state === 'empty'", script);
            Assert.Contains("state === 'error'", script);
            Assert.Contains("state !== 'loading'", script);
        }

        private static string ReadWebUiFile(string fileName)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(
                    current.FullName,
                    "OutlookAI",
                    "WebUI",
                    fileName);
                if (File.Exists(candidate))
                {
                    return File.ReadAllText(candidate);
                }
                current = current.Parent;
            }

            throw new FileNotFoundException("Could not find WebUI/" + fileName);
        }
    }
}
