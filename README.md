# OutlookAI

OutlookAI is a VSTO add-in for Microsoft Outlook that adds AI-assisted email drafting, editing, mailbox copilot chat, reports, and exports.

This branch uses a configurable **LiteLLM** connector. The add-in talks to an OpenAI-compatible LiteLLM proxy and users enter their own API key in OutlookAI Settings.

## LiteLLM Configuration

Install-time scripts provide the shared defaults:

- `LiteLlmBaseUrl`
- `Model`
- `VoiceModel`
- `Temperature`
- `MaxTokens`
- `MaxBulkExportRows`

Users provide only:

- `LiteLlmApiKey`

The API key is stored per user in `%APPDATA%\OutlookAI\config.xml`. It is not written to the machine-wide install config.

The runtime endpoints are:

- `<LiteLlmBaseUrl>/chat/completions`
- `<LiteLlmBaseUrl>/audio/transcriptions`

## Install

Publish the VSTO build, then run the installer as Administrator:

```powershell
.\Deploy\Install-OutlookAI.ps1 `
  -SourcePath "C:\OutlookAI" `
  -LiteLlmBaseUrl "https://litellm.company.example/v1" `
  -LiteLlmModel "company/outlook-chat" `
  -LiteLlmVoiceModel "company/outlook-transcribe" `
  -Temperature 0.2 `
  -MaxTokens 4096
```

After install, each user opens OutlookAI Settings, enters the admin password, and saves their LiteLLM API key.

## Project Layout

- `VSTO2/OutlookAI/Services/LiteLlmChatService.cs` - LiteLLM chat completions connector, streaming, and tool-call loop.
- `VSTO2/OutlookAI/Services/LiteLlmCredentialService.cs` - per-user API key state.
- `VSTO2/OutlookAI/Services/LiteLlmVoiceService.cs` - LiteLLM audio transcription connector.
- `VSTO2/OutlookAI/Services/Tools/` - Outlook tool surface used by chat and reports.
- `Deploy/` - installer, uninstaller, and deployment docs.

## Development

The project targets .NET Framework 4.7.2 and VSTO. Building the add-in requires Visual Studio/MSBuild with Office/VSTO targets installed.

Tests:

```powershell
dotnet restore VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
dotnet test VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
```
