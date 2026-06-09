# Install

Use `Deploy/Install-OutlookAI.ps1` for workstation, RDS, and silent image installs.

The installer writes the server-side LiteLLM defaults to:

`C:\Program Files\OutlookAI\config.xml`

Example:

```powershell
.\Deploy\Install-OutlookAI.ps1 `
  -SourcePath "C:\OutlookAI" `
  -LiteLlmBaseUrl "https://litellm.company.example/v1" `
  -LiteLlmModel "company/outlook-chat" `
  -LiteLlmVoiceModel "company/outlook-transcribe" `
  -Temperature 0.2 `
  -MaxTokens 4096
```

The installer does **not** write API keys. Each user opens OutlookAI Settings and enters their own LiteLLM API key, which is stored in that user's `%APPDATA%\OutlookAI\config.xml`.

Basic verification:

1. Outlook shows the `AI Assistant` ribbon group.
2. Open Settings and enter the admin password.
3. Confirm the LiteLLM endpoint/model values are shown.
4. Enter the user's LiteLLM API key.
5. Run a quick action or send a chat message.
