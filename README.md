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

`VoiceModel` необязателен. Если модель транскрибации не используется, оставьте значение пустым или укажите `null`; голосовая транскрибация будет отключена.

Пользователь указывает только:

- `LiteLlmApiKey`

Ключ API хранится отдельно для каждого пользователя в `%APPDATA%\OutlookAI\config.xml`. Он не записывается в общую машинную конфигурацию установки.

Используемые конечные точки:

- `<LiteLlmBaseUrl>/chat/completions`
- `<LiteLlmBaseUrl>/audio/transcriptions`

## Установка

Для локального запуска без командной строки соберите единый EXE-установщик:

```powershell
.\Deploy\Make-InstallerExe.ps1 `
  -Tag v3.0.0 `
  -LiteLlmBaseUrl "https://litellm.company.example/v1" `
  -LiteLlmModel "company/outlook-chat" `
  -LiteLlmVoiceModel "" `
  -Temperature 0.2 `
  -MaxTokens 4096
```

Скрипт создаёт `out\OutlookAI-v3.0.0-Setup.exe`. Для сборки нужен MSBuild/Visual Studio с VSTO targets и встроенный Windows `iexpress.exe`; конечному пользователю Visual Studio для запуска готового EXE не нужна. Пользователь запускает этот файл двойным кликом; установщик сам запросит права администратора через UAC.

Для ручной или тихой установки можно опубликовать VSTO-сборку и запустить скрипт от имени администратора:

```powershell
.\Deploy\Install-OutlookAI.ps1 `
  -SourcePath "C:\OutlookAI" `
  -LiteLlmBaseUrl "https://litellm.company.example/v1" `
  -LiteLlmModel "company/outlook-chat" `
  -LiteLlmVoiceModel "" `
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
