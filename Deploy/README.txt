OutlookAI - Deployment Guide
============================

OutlookAI v3 uses a configurable LiteLLM connector.

The installer writes server defaults to:
  C:\Program Files\OutlookAI\config.xml

Users provide their own LiteLLM API key in OutlookAI Settings. API keys are
stored per user under:
  %APPDATA%\OutlookAI\config.xml


PREREQUISITES
-------------
- Windows 10 / 11 or Windows Server 2019 / 2022 / 2025.
- Microsoft Outlook desktop.
- .NET Framework 4.7.2 or later.
- Visual Studio Tools for Office Runtime:
  https://aka.ms/VSTORuntime
- Microsoft Edge WebView2 Evergreen Runtime.


INSTALL
-------
Publish the VSTO build, then run PowerShell as Administrator:

  .\Deploy\Install-OutlookAI.ps1 `
    -SourcePath "C:\OutlookAI" `
    -LiteLlmBaseUrl "https://litellm.company.example/v1" `
    -LiteLlmModel "company/outlook-chat" `
    -LiteLlmVoiceModel "company/outlook-transcribe" `
    -Temperature 0.2 `
    -MaxTokens 4096

The LiteLLM base URL should be the OpenAI-compatible API root. The add-in
will call:
  <LiteLlmBaseUrl>/chat/completions
  <LiteLlmBaseUrl>/audio/transcriptions


WHAT THE INSTALLER DOES
-----------------------
1. Cleans stale OutlookAI VSTO/ClickOnce registrations.
2. Copies the published build to C:\Program Files\OutlookAI.
3. Writes config.xml with LiteLLM base URL, model, voice model, temperature,
   max tokens, and max bulk export rows.
4. Installs or verifies WebView2.
5. Registers the Outlook add-in for all users.

The installer does not install, share, or rotate API keys.


USER FIRST RUN
--------------
1. Open Outlook.
2. Open OutlookAI Settings.
3. Enter the admin password.
4. Enter the user's LiteLLM API key.
5. Run a quick action or send a chat message.


UNINSTALL
---------
Run PowerShell as Administrator:

  .\Deploy\Uninstall-OutlookAI.ps1

This removes the HKLM Outlook add-in registration and
C:\Program Files\OutlookAI. Per-user API keys remain in each user's AppData
unless their %APPDATA%\OutlookAI\config.xml is deleted.
