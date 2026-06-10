using System;
using System.IO;
using OutlookAI;
using Xunit;

namespace OutlookAI.Tests
{
    [Collection("Config")]
    public class ConfigTests
    {
        private static (string global, string user) MakeTempPaths()
        {
            var dir = Path.Combine(Path.GetTempPath(),
                "outlookai-config-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(dir);
            return (Path.Combine(dir, "global.xml"), Path.Combine(dir, "user.xml"));
        }

        [Fact]
        public void LoadConfigFromPaths_UsesLiteLlmDefaults_WhenFilesAreMissing()
        {
            var (g, u) = MakeTempPaths();
            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("https://litellm.example.com/v1", Config.LiteLlmBaseUrl);
            Assert.Equal("", Config.LiteLlmApiKey);
            Assert.Equal("gpt-4.1-mini", Config.Model);
            Assert.Equal("", Config.VoiceModel);
            Assert.Equal(0.2, Config.Temperature);
            Assert.Equal(4096, Config.MaxTokens);
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesServerLiteLlmDefaultsFromGlobal()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<LiteLlmBaseUrl>https://llm.example.test/v1/</LiteLlmBaseUrl>"
                + "<Model>company/chat</Model>"
                + "<VoiceModel>company/transcribe</VoiceModel>"
                + "<Temperature>0.7</Temperature>"
                + "<MaxTokens>8192</MaxTokens>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("https://llm.example.test/v1", Config.LiteLlmBaseUrl);
            Assert.Equal("company/chat", Config.Model);
            Assert.Equal("company/transcribe", Config.VoiceModel);
            Assert.Equal(0.7, Config.Temperature);
            Assert.Equal(8192, Config.MaxTokens);
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesServerDefaultsFromAnyMachineConfigPath()
        {
            var dir = Path.Combine(Path.GetTempPath(),
                "outlookai-config-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(dir);
            var missing32BitPath = Path.Combine(dir, "missing-x86.xml");
            var installedPath = Path.Combine(dir, "program-files.xml");
            var userPath = Path.Combine(dir, "user.xml");

            File.WriteAllText(installedPath, "<Config>"
                + "<LiteLlmBaseUrl>http://localhost:11434</LiteLlmBaseUrl>"
                + "<Model>ollama/qwen2.5:3b</Model>"
                + "</Config>");

            Config.LoadConfigFromPaths(new[] { missing32BitPath, installedPath }, null, userPath);

            Assert.Equal("http://localhost:11434/v1", Config.LiteLlmBaseUrl);
            Assert.Equal("ollama/qwen2.5:3b", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesServerDefaultsFromSharedConfig()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(s, "<Config>"
                + "<LiteLlmBaseUrl>http://localhost:11434</LiteLlmBaseUrl>"
                + "<Model>ollama/qwen2.5:3b</Model>"
                + "</Config>");

            try
            {
                Config.LoadConfigFromPaths(g, s, u);

                Assert.Equal("http://localhost:11434/v1", Config.LiteLlmBaseUrl);
                Assert.Equal("ollama/qwen2.5:3b", Config.Model);
            }
            finally
            {
                if (File.Exists(s)) File.Delete(s);
            }
        }

        [Theory]
        [InlineData("http://localhost:11434", "http://localhost:11434/v1")]
        [InlineData("http://localhost:11434/", "http://localhost:11434/v1")]
        [InlineData("https://llm.example.test/v1/", "https://llm.example.test/v1")]
        [InlineData("https://llm.example.test/v1/chat/completions", "https://llm.example.test/v1")]
        public void NormalizeBaseUrl_AcceptsHostOrOpenAiEndpoint(string raw, string expected)
        {
            Assert.Equal(expected, Config.NormalizeBaseUrl(raw));
        }

        [Fact]
        public void LoadConfigFromPaths_EmptyVoiceModelDisablesTranscription()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config><VoiceModel></VoiceModel></Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("", Config.VoiceModel);
        }

        [Fact]
        public void LoadConfigFromPaths_UserApiKeyOverridesOnlyApiKey()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<LiteLlmBaseUrl>https://llm.example.test/v1</LiteLlmBaseUrl>"
                + "<Model>server-model</Model>"
                + "</Config>");
            File.WriteAllText(u, "<Config>"
                + "<LiteLlmApiKey>sk-user</LiteLlmApiKey>"
                + "<LiteLlmBaseUrl>https://malicious.example/v1</LiteLlmBaseUrl>"
                + "<Model>user-model</Model>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("sk-user", Config.LiteLlmApiKey);
            Assert.Equal("https://llm.example.test/v1", Config.LiteLlmBaseUrl);
            Assert.Equal("server-model", Config.Model);
        }

        [Fact]
        public void LoadConfigFromPaths_DoesNotReadApiKeyFromGlobalOrShared()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var s = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");

            File.WriteAllText(g, "<Config><LiteLlmApiKey>sk-global</LiteLlmApiKey></Config>");
            File.WriteAllText(s, "<Config><LiteLlmApiKey>sk-shared</LiteLlmApiKey></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, s, u);
                Assert.Equal("", Config.LiteLlmApiKey);
            }
            finally
            {
                if (File.Exists(g)) File.Delete(g);
                if (File.Exists(s)) File.Delete(s);
            }
        }

        [Fact]
        public void LoadConfigFromPaths_UserOverridesReasoningEffortAndWriteTools()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<ReasoningEffort>Medium</ReasoningEffort>"
                + "<WriteToolsEnabled>true</WriteToolsEnabled>"
                + "</Config>");
            File.WriteAllText(u, "<Config>"
                + "<ReasoningEffort>Low</ReasoningEffort>"
                + "<WriteToolsEnabled>false</WriteToolsEnabled>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal("Low", Config.ReasoningEffort);
            Assert.False(Config.WriteToolsEnabled);
        }

        [Fact]
        public void AvailableReasoningEfforts_ContainsGenericOptions()
        {
            var expected = new[] { "None", "Minimal", "Low", "Medium", "High", "XHigh" };
            Assert.Equal(expected, Config.AvailableReasoningEfforts);
            Assert.Equal(expected, Config.ReasoningEffortsForModel("any-model"));
        }

        [Fact]
        public void LoadConfigFromPaths_AppliesEnabledWriteToolsFromCSV()
        {
            var (g, u) = MakeTempPaths();
            File.WriteAllText(g, "<Config>"
                + "<EnabledWriteTools>outlook_create_draft, outlook_set_category</EnabledWriteTools>"
                + "</Config>");

            Config.LoadConfigFromPaths(g, u);

            Assert.Equal(2, Config.EnabledWriteTools.Count);
            Assert.Contains("outlook_create_draft", Config.EnabledWriteTools);
            Assert.Contains("outlook_set_category", Config.EnabledWriteTools);
            Assert.DoesNotContain("outlook_mark_as_read", Config.EnabledWriteTools);
        }

        [Fact]
        public void MaxBulkExportRows_LoadsFromGlobalConfig_AndClamps()
        {
            var g = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            var u = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(g, "<Config><MaxBulkExportRows>999999</MaxBulkExportRows></Config>");
            try
            {
                Config.LoadConfigFromPaths(g, sharedConfigPath: null, userConfigPath: u);
                Assert.Equal(10000, Config.MaxBulkExportRows);
            }
            finally { if (File.Exists(g)) File.Delete(g); }
        }
    }
}
