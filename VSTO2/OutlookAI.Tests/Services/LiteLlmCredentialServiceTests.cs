using OutlookAI.Services;
using Xunit;

namespace OutlookAI.Tests.Services
{
    [Collection("Config")]
    public class LiteLlmCredentialServiceTests
    {
        [Fact]
        public void GetStatus_ReturnsMissing_WhenApiKeyAbsent()
        {
            Config.ResetDefaults();
            using (var svc = new LiteLlmCredentialService(() => { }))
            {
                Assert.Equal(CredentialState.Missing, svc.GetStatus().State);
            }
        }

        [Fact]
        public void SaveApiKey_PersistsToConfig_AndReportsConfigured()
        {
            Config.ResetDefaults();
            using (var svc = new LiteLlmCredentialService(() => { }))
            {
                svc.SaveApiKey(" sk-test ");

                Assert.Equal("sk-test", Config.LiteLlmApiKey);
                Assert.Equal(CredentialState.Configured, svc.GetStatus().State);
            }
        }

        [Fact]
        public void GetApiKey_Throws_WhenMissing()
        {
            Config.ResetDefaults();
            using (var svc = new LiteLlmCredentialService(() => { }))
            {
                Assert.Throws<System.InvalidOperationException>(() => svc.GetApiKey());
            }
        }
    }
}
