using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OutlookAI.Diagnostics;

namespace OutlookAI
{
    public static class Config
    {
        // ============================================================
        // CONFIGURATION DEFAULTS (v3 - LiteLLM API-key connector)
        //
        // Install scripts own the server defaults: LiteLlmBaseUrl, models,
        // temperature, token limits, and bulk export limits. Users provide
        // only their own LiteLLM API key in Settings; it is stored in their
        // AppData config and is never written to the shared server config.
        // ============================================================

        public const string DefaultLiteLlmBaseUrl = "https://litellm.example.com/v1";
        public const string DefaultModel = "gpt-4.1-mini";
        public const string DefaultVoiceModel = "";
        public const string DefaultReasoningEffort = "None";
        public const double DefaultTemperature = 0.2;
        public const int DefaultMaxTokens = 4096;
        public const bool DefaultWriteToolsEnabled = true;
        public const int DefaultMaxBulkExportRows = 2000;

        private const int MinBulkExportRows = 1;
        private const int MaxBulkExportRowsCeiling = Services.Tools.BulkExportRowCap.Max;
        private const int MinMaxTokens = 1;
        private const int MaxMaxTokens = 200000;

        public static string LiteLlmBaseUrl { get; set; } = DefaultLiteLlmBaseUrl;
        public static string LiteLlmApiKey { get; set; } = "";
        public static string Model { get; set; } = DefaultModel;
        public static string VoiceModel { get; set; } = DefaultVoiceModel;
        public static string ReasoningEffort { get; set; } = DefaultReasoningEffort;
        public static double Temperature { get; set; } = DefaultTemperature;
        public static int MaxTokens { get; set; } = DefaultMaxTokens;
        public static bool WriteToolsEnabled { get; set; } = DefaultWriteToolsEnabled;
        public static int MaxBulkExportRows { get; set; } = DefaultMaxBulkExportRows;

        public static readonly string[] AllWriteTools =
        {
            "outlook_create_draft",
            "outlook_mark_as_read",
            "outlook_flag_message",
            "outlook_set_category"
        };

        public static HashSet<string> EnabledWriteTools { get; set; } =
            new HashSet<string>(AllWriteTools, StringComparer.Ordinal);

        public static readonly string[] AvailableReasoningEfforts =
        {
            "None",
            "Minimal",
            "Low",
            "Medium",
            "High",
            "XHigh"
        };

        public static string[] ReasoningEffortsForModel(string model)
        {
            return AvailableReasoningEfforts;
        }

        public static bool IsUsingPlaceholderLiteLlmEndpoint()
        {
            return string.Equals(
                NormalizeBaseUrl(LiteLlmBaseUrl),
                NormalizeBaseUrl(DefaultLiteLlmBaseUrl),
                StringComparison.OrdinalIgnoreCase);
        }

        public static void EnsureLiteLlmServerConfigured()
        {
            if (IsUsingPlaceholderLiteLlmEndpoint())
            {
                throw new InvalidOperationException(
                    "Серверная конфигурация LiteLLM не установлена. "
                    + "Переустановите OutlookAI через installer и укажите LiteLLM base URL, например http://localhost:4000/v1, и модель, например local-model.");
            }
        }

        private static readonly string[] GlobalConfigFilePaths = BuildGlobalConfigFilePaths();

        private static readonly string UserConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OutlookAI",
            "config.xml"
        );

        private static readonly string SharedConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "OutlookAI",
            "config.xml"
        );

        static Config()
        {
            LoadConfig();
        }

        public static void LoadConfig()
        {
            LoadConfigFromPaths(GlobalConfigFilePaths, SharedConfigFilePath, UserConfigFilePath);
        }

        public static void LoadConfigFromPaths(string globalConfigPath, string sharedConfigPath, string userConfigPath)
        {
            LoadConfigFromPaths(new[] { globalConfigPath }, sharedConfigPath, userConfigPath);
        }

        public static void LoadConfigFromPaths(IEnumerable<string> globalConfigPaths, string sharedConfigPath, string userConfigPath)
        {
            ResetDefaults();
            foreach (var globalConfigPath in globalConfigPaths ?? Enumerable.Empty<string>())
            {
                LoadFromFile(globalConfigPath, allowServerFields: true, allowApiKey: false);
            }
            LoadFromFile(sharedConfigPath, allowServerFields: true, allowApiKey: false);
            LoadFromFile(userConfigPath, allowServerFields: false, allowApiKey: true);
            TraceEffectiveConfig(globalConfigPaths, sharedConfigPath, userConfigPath);
        }

        public static void LoadConfigFromPaths(string globalConfigPath, string userConfigPath)
        {
            LoadConfigFromPaths(globalConfigPath, sharedConfigPath: null, userConfigPath);
        }

        private static string[] BuildGlobalConfigFilePaths()
        {
            var paths = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddConfigPath(string root)
            {
                if (string.IsNullOrWhiteSpace(root))
                {
                    return;
                }

                try
                {
                    var path = Path.Combine(root, "OutlookAI", "config.xml");
                    if (seen.Add(path))
                    {
                        paths.Add(path);
                    }
                }
                catch
                {
                    // Ignore malformed environment paths.
                }
            }

            AddConfigPath(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86));
            AddConfigPath(Environment.GetEnvironmentVariable("ProgramFiles(x86)"));
            AddConfigPath(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles));
            AddConfigPath(Environment.GetEnvironmentVariable("ProgramW6432"));
            AddConfigPath(Environment.GetEnvironmentVariable("ProgramFiles"));

            return paths.ToArray();
        }

        public static void ResetDefaults()
        {
            LiteLlmBaseUrl = DefaultLiteLlmBaseUrl;
            LiteLlmApiKey = "";
            Model = DefaultModel;
            VoiceModel = DefaultVoiceModel;
            ReasoningEffort = DefaultReasoningEffort;
            Temperature = DefaultTemperature;
            MaxTokens = DefaultMaxTokens;
            WriteToolsEnabled = DefaultWriteToolsEnabled;
            MaxBulkExportRows = DefaultMaxBulkExportRows;
            EnabledWriteTools = new HashSet<string>(AllWriteTools, StringComparer.Ordinal);
        }

        private static void LoadFromFile(string filePath, bool allowServerFields, bool allowApiKey)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    return;
                }

                var doc = XDocument.Load(filePath);
                var root = doc.Root;
                if (root == null)
                {
                    return;
                }

                var reasoningEffort = root.Element("ReasoningEffort");
                if (reasoningEffort != null && !string.IsNullOrWhiteSpace(reasoningEffort.Value))
                {
                    foreach (var allowed in AvailableReasoningEfforts)
                    {
                        if (string.Equals(allowed, reasoningEffort.Value, StringComparison.OrdinalIgnoreCase))
                        {
                            ReasoningEffort = allowed;
                            break;
                        }
                    }
                }

                var writeToolsEnabled = root.Element("WriteToolsEnabled");
                if (writeToolsEnabled != null && bool.TryParse(writeToolsEnabled.Value, out var wte))
                {
                    WriteToolsEnabled = wte;
                }

                var enabledWriteTools = root.Element("EnabledWriteTools");
                if (enabledWriteTools != null && !string.IsNullOrWhiteSpace(enabledWriteTools.Value))
                {
                    var requested = enabledWriteTools.Value
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(x => x.Trim());
                    var canonical = new HashSet<string>(AllWriteTools, StringComparer.Ordinal);
                    EnabledWriteTools = new HashSet<string>(
                        requested.Where(canonical.Contains),
                        StringComparer.Ordinal);
                }

                if (allowApiKey)
                {
                    var apiKey = root.Element("LiteLlmApiKey") ?? root.Element("ApiKey");
                    if (apiKey != null)
                    {
                        LiteLlmApiKey = apiKey.Value ?? "";
                    }
                }

                if (!allowServerFields)
                {
                    return;
                }

                var baseUrl = root.Element("LiteLlmBaseUrl") ?? root.Element("BaseUrl");
                if (baseUrl != null && !string.IsNullOrWhiteSpace(baseUrl.Value))
                {
                    LiteLlmBaseUrl = NormalizeBaseUrl(baseUrl.Value);
                }

                var model = root.Element("Model") ?? root.Element("LiteLlmModel");
                if (model != null && !string.IsNullOrWhiteSpace(model.Value))
                {
                    Model = model.Value.Trim();
                }

                var voiceModel = root.Element("VoiceModel") ?? root.Element("LiteLlmVoiceModel");
                if (voiceModel != null)
                {
                    VoiceModel = voiceModel.Value.Trim();
                }

                var temperature = root.Element("Temperature");
                if (temperature != null && double.TryParse(
                    temperature.Value,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out var t))
                {
                    if (t < 0) t = 0;
                    if (t > 2) t = 2;
                    Temperature = t;
                }

                var maxTokens = root.Element("MaxTokens");
                if (maxTokens != null && int.TryParse(maxTokens.Value, out var mt))
                {
                    if (mt < MinMaxTokens) mt = MinMaxTokens;
                    if (mt > MaxMaxTokens) mt = MaxMaxTokens;
                    MaxTokens = mt;
                }

                var maxBulkExportRows = root.Element("MaxBulkExportRows");
                if (maxBulkExportRows != null && int.TryParse(maxBulkExportRows.Value, out var mber))
                {
                    if (mber < MinBulkExportRows) mber = MinBulkExportRows;
                    if (mber > MaxBulkExportRowsCeiling) mber = MaxBulkExportRowsCeiling;
                    MaxBulkExportRows = mber;
                }
            }
            catch
            {
                // Skip if file is missing or invalid; defaults stay in place.
            }
        }

        private static void TraceEffectiveConfig(IEnumerable<string> globalConfigPaths, string sharedConfigPath, string userConfigPath)
        {
            try
            {
                var machinePaths = string.Join("; ",
                    (globalConfigPaths ?? Enumerable.Empty<string>())
                        .Where(path => !string.IsNullOrWhiteSpace(path))
                        .Select(path => path + (File.Exists(path) ? " [found]" : " [missing]")));
                var sharedState = string.IsNullOrWhiteSpace(sharedConfigPath)
                    ? "<none>"
                    : sharedConfigPath + (File.Exists(sharedConfigPath) ? " [found]" : " [missing]");
                var userState = string.IsNullOrWhiteSpace(userConfigPath)
                    ? "<none>"
                    : userConfigPath + (File.Exists(userConfigPath) ? " [found]" : " [missing]");

                TraceLog.Write(
                    "Config loaded. MachinePaths=" + machinePaths
                    + "; SharedPath=" + sharedState
                    + "; UserPath=" + userState
                    + "; LiteLlmBaseUrl=" + NormalizeBaseUrl(LiteLlmBaseUrl)
                    + "; Model=" + Model
                    + "; VoiceModel=" + (string.IsNullOrWhiteSpace(VoiceModel) ? "<disabled>" : VoiceModel)
                    + "; ApiKeyConfigured=" + (!string.IsNullOrWhiteSpace(LiteLlmApiKey)),
                    "Config");
            }
            catch
            {
                // Diagnostics must not affect startup.
            }
        }

        public static void SaveConfig()
        {
            var userDoc = new XDocument(
                new XElement("Config",
                    new XElement("LiteLlmApiKey", LiteLlmApiKey ?? ""),
                    new XElement("ReasoningEffort", ReasoningEffort),
                    new XElement("WriteToolsEnabled", WriteToolsEnabled),
                    new XElement("EnabledWriteTools",
                        string.Join(",", EnabledWriteTools ?? new HashSet<string>()))
                )
            );

            TrySaveTo(UserConfigFilePath, userDoc);
        }

        public static string NormalizeBaseUrl(string raw)
        {
            var value = (raw ?? "").Trim();
            while (value.EndsWith("/", StringComparison.Ordinal))
            {
                value = value.Substring(0, value.Length - 1);
            }
            if (string.IsNullOrEmpty(value))
            {
                return DefaultLiteLlmBaseUrl;
            }

            if (value.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                value = value.Substring(0, value.Length - "/chat/completions".Length);
            }
            else if (value.EndsWith("/audio/transcriptions", StringComparison.OrdinalIgnoreCase))
            {
                value = value.Substring(0, value.Length - "/audio/transcriptions".Length);
            }

            while (value.EndsWith("/", StringComparison.Ordinal))
            {
                value = value.Substring(0, value.Length - 1);
            }

            if (!value.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
            {
                value += "/v1";
            }

            return value;
        }

        private static void TrySaveTo(string filePath, XDocument doc)
        {
            try
            {
                var dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                doc.Save(filePath);
            }
            catch (Exception ex)
            {
                try
                {
                    OutlookAI.Diagnostics.TraceLog.Write(
                        "Config.TrySaveTo failed for '" + filePath + "': " + ex.Message,
                        "Config");
                }
                catch { }
            }
        }
    }
}
