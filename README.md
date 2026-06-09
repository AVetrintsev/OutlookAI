# OutlookAI

OutlookAI - VSTO-надстройка для Microsoft Outlook. Она добавляет AI-помощника для написания и редактирования писем, чат по почтовому ящику, отчёты и экспорт данных.

Текущая ветка использует настраиваемый **LiteLLM**-коннектор. Надстройка обращается к LiteLLM-прокси с OpenAI-compatible API, а каждый пользователь указывает собственный ключ API в настройках OutlookAI.

## Конфигурация LiteLLM

Базовые параметры задаются при установке расширения скриптом:

- `LiteLlmBaseUrl`
- `Model`
- `VoiceModel`
- `Temperature`
- `MaxTokens`
- `MaxBulkExportRows`

Пользователь указывает только:

- `LiteLlmApiKey`

Ключ API хранится отдельно для каждого пользователя в `%APPDATA%\OutlookAI\config.xml`. Он не записывается в общую машинную конфигурацию установки.

Используемые конечные точки:

- `<LiteLlmBaseUrl>/chat/completions`
- `<LiteLlmBaseUrl>/audio/transcriptions`

## Установка

Опубликуйте VSTO-сборку, затем запустите установщик от имени администратора:

```powershell
.\Deploy\Install-OutlookAI.ps1 `
  -SourcePath "C:\OutlookAI" `
  -LiteLlmBaseUrl "https://litellm.company.example/v1" `
  -LiteLlmModel "company/outlook-chat" `
  -LiteLlmVoiceModel "company/outlook-transcribe" `
  -Temperature 0.2 `
  -MaxTokens 4096
```

После установки каждый пользователь открывает настройки OutlookAI и сохраняет свой ключ API LiteLLM.

## Структура проекта

- `VSTO2/OutlookAI/Services/LiteLlmChatService.cs` - коннектор LiteLLM chat completions, потоковая выдача и цикл tool-calling.
- `VSTO2/OutlookAI/Services/LiteLlmCredentialService.cs` - состояние пользовательского ключа API.
- `VSTO2/OutlookAI/Services/LiteLlmVoiceService.cs` - коннектор LiteLLM audio transcription.
- `VSTO2/OutlookAI/Services/Tools/` - слой Outlook-инструментов, который используется чатом и отчётами.
- `Deploy/` - установщик, скрипт удаления и инструкции по развёртыванию.

## Разработка

Проект использует .NET Framework 4.7.2 и VSTO. Для сборки надстройки нужен Visual Studio/MSBuild с установленными Office/VSTO targets.

Тесты:

```powershell
dotnet restore VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
dotnet test VSTO2\OutlookAI.Tests\OutlookAI.Tests.csproj
```
