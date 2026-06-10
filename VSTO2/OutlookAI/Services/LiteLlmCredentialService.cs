using System;
using System.Threading.Tasks;

namespace OutlookAI.Services
{
    public enum CredentialState
    {
        Missing,
        Configured,
        Error
    }

    public sealed class CredentialStatus
    {
        public CredentialState State { get; }
        public string Message { get; }

        public CredentialStatus(CredentialState state, string message)
        {
            State = state;
            Message = message ?? "";
        }

        public static CredentialStatus Missing(string message = "Ключ API LiteLLM не настроен")
            => new CredentialStatus(CredentialState.Missing, message);

        public static CredentialStatus Configured()
            => new CredentialStatus(CredentialState.Configured, "Ключ API LiteLLM настроен");

        public static CredentialStatus Error(string message)
            => new CredentialStatus(CredentialState.Error, message);
    }

    public sealed class LiteLlmCredentialService : IDisposable
    {
        private readonly Action _saveConfig;

        public event EventHandler<CredentialStatus> StatusChanged;

        public LiteLlmCredentialService()
            : this(Config.SaveConfig)
        {
        }

        public LiteLlmCredentialService(Action saveConfig)
        {
            _saveConfig = saveConfig ?? Config.SaveConfig;
        }

        public string GetApiKey()
        {
            var apiKey = Config.LiteLlmApiKey ?? "";
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                throw new InvalidOperationException("Ключ API LiteLLM не настроен. Откройте настройки и введите ключ API.");
            }
            return apiKey.Trim();
        }

        public CredentialStatus GetStatus()
        {
            try
            {
                return string.IsNullOrWhiteSpace(Config.LiteLlmApiKey)
                    ? CredentialStatus.Missing()
                    : CredentialStatus.Configured();
            }
            catch (Exception ex)
            {
                return CredentialStatus.Error(ex.Message);
            }
        }

        public void SaveApiKey(string apiKey)
        {
            Config.LiteLlmApiKey = (apiKey ?? "").Trim();
            _saveConfig();
            RaiseStatusChanged(GetStatus());
        }

        public Task ClearApiKeyAsync()
        {
            SaveApiKey("");
            return Task.FromResult(0);
        }

        private void RaiseStatusChanged(CredentialStatus status)
        {
            var handler = StatusChanged;
            if (handler == null)
            {
                return;
            }
            try
            {
                handler(this, status);
            }
            catch
            {
                // Subscriber bugs must not break credential state.
            }
        }

        public void Dispose()
        {
        }
    }
}
