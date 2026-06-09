# Установка

Используйте `Deploy/Install-OutlookAI.ps1` для установки на рабочую станцию, RDS/Terminal Server или в тихий образ развёртывания.

Установщик записывает серверные значения LiteLLM по умолчанию в:

`C:\Program Files\OutlookAI\config.xml`

Пример:

```powershell
.\Deploy\Install-OutlookAI.ps1 `
  -SourcePath "C:\OutlookAI" `
  -LiteLlmBaseUrl "https://litellm.company.example/v1" `
  -LiteLlmModel "company/outlook-chat" `
  -LiteLlmVoiceModel "company/outlook-transcribe" `
  -Temperature 0.2 `
  -MaxTokens 4096
```

Установщик **не** записывает ключи API. Каждый пользователь открывает настройки OutlookAI и вводит собственный ключ API LiteLLM. Ключ хранится в пользовательском `%APPDATA%\OutlookAI\config.xml`.

Пароль администратора для ввода пользовательского ключа API не нужен. Он должен использоваться только для административных настроек.

Базовая проверка:

1. В Outlook отображается группа ленты `AI Assistant`.
2. Откройте настройки OutlookAI.
3. Убедитесь, что отображаются правильные значения LiteLLM endpoint/model.
4. Введите пользовательский ключ API LiteLLM.
5. Запустите быструю команду или отправьте сообщение в чат.
