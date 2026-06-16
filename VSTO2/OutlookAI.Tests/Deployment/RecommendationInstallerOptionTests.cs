using System.IO;
using Xunit;

namespace OutlookAI.Tests.Deployment
{
    public sealed class RecommendationInstallerOptionTests
    {
        [Fact]
        public void InstallerPipeline_CarriesMachineRecommendationSetting()
        {
            var install = ReadRepoFile("Deploy", "Install-OutlookAI.ps1");
            var builder = ReadRepoFile("Deploy", "Make-InstallerExe.ps1");
            var setup = ReadRepoFile("Deploy", "OutlookAI-Setup.ps1");
            var interactive = ReadRepoFile("Deploy", "Build-InstallerInteractive.ps1");

            Assert.Contains("[string]$RecommendationsEnabled = \"false\"", install);
            Assert.Contains("ConvertTo-BooleanFlag", install);
            Assert.Contains("<RecommendationsEnabled>", install);
            Assert.Contains("RecommendationsEnabled = $RecommendationsEnabled", builder);
            Assert.Contains("Add-BooleanInstallArgument", setup);
            Assert.Contains("$Arguments.Add($Name)", setup);
            Assert.Contains("\"1\"", setup);
            Assert.Contains("\"0\"", setup);
            Assert.Contains("Enable AI recommendations", interactive);
            Assert.Contains("RecommendationsEnabled = $recommendationsEnabled", interactive);
        }

        private static string ReadRepoFile(params string[] parts)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, Path.Combine(parts));
                if (File.Exists(candidate))
                {
                    return File.ReadAllText(candidate);
                }
                current = current.Parent;
            }

            throw new FileNotFoundException("Could not find " + Path.Combine(parts));
        }
    }
}
