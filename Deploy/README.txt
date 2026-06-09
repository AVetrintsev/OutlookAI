OutlookAI - инструкция по развёртыванию
======================================

OutlookAI v3 использует настраиваемый LiteLLM-коннектор.

Установщик записывает серверные значения по умолчанию в:
  C:\Program Files\OutlookAI\config.xml

Пользователи вводят собственный ключ API LiteLLM в настройках OutlookAI.
Ключи API хранятся отдельно для каждого пользователя:
  %APPDATA%\OutlookAI\config.xml

Пароль администратора для ввода пользовательского ключа API не нужен.


ТРЕБОВАНИЯ
----------
- Windows 10 / 11 или Windows Server 2019 / 2022 / 2025.
- Microsoft Outlook для рабочего стола.
- .NET Framework 4.7.2 или новее.
- Visual Studio Tools for Office Runtime:
  https://aka.ms/VSTORuntime
- Microsoft Edge WebView2 Evergreen Runtime.


УСТАНОВКА
---------
Опубликуйте VSTO-сборку, затем запустите PowerShell от имени администратора:

  .\Deploy\Install-OutlookAI.ps1 `
    -SourcePath "C:\OutlookAI" `
    -LiteLlmBaseUrl "https://litellm.company.example/v1" `
    -LiteLlmModel "company/outlook-chat" `
    -LiteLlmVoiceModel "company/outlook-transcribe" `
    -Temperature 0.2 `
    -MaxTokens 4096

`LiteLlmBaseUrl` должен указывать на корень OpenAI-compatible API.
Надстройка будет вызывать:
  <LiteLlmBaseUrl>/chat/completions
  <LiteLlmBaseUrl>/audio/transcriptions


ЧТО ДЕЛАЕТ УСТАНОВЩИК
--------------------
1. Очищает старые регистрации OutlookAI VSTO/ClickOnce.
2. Копирует опубликованную сборку в C:\Program Files\OutlookAI.
3. Записывает config.xml с LiteLLM base URL, model, voice model,
   temperature, max tokens и max bulk export rows.
4. Устанавливает или проверяет WebView2.
5. Регистрирует надстройку Outlook для всех пользователей.

Установщик не устанавливает, не расшаривает и не ротирует ключи API.


ПЕРВЫЙ ЗАПУСК ПОЛЬЗОВАТЕЛЯ
--------------------------
1. Откройте Outlook.
2. Откройте настройки OutlookAI.
3. Введите ключ API LiteLLM пользователя.
4. Запустите быструю команду или отправьте сообщение в чат.

Пароль администратора нужен только для административных настроек, а не для
ввода пользовательского ключа API.


УДАЛЕНИЕ
--------
Запустите PowerShell от имени администратора:

  .\Deploy\Uninstall-OutlookAI.ps1

Скрипт удаляет регистрацию надстройки Outlook в HKLM и
C:\Program Files\OutlookAI. Пользовательские ключи API остаются в AppData,
пока не будет удалён файл %APPDATA%\OutlookAI\config.xml для конкретного
пользователя.
