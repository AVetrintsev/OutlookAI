using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

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
        public const string DefaultVoiceModel = "gpt-4o-mini-transcribe";
        public const string DefaultReasoningEffort = "None";
        public const double DefaultTemperature = 0.2;
        public const int DefaultMaxTokens = 4096;
        public const bool DefaultWriteToolsEnabled = true;
        public const int DefaultMaxBulkExportRows = 2000;

        private const int MinBulkExportRows = 1;
        private const int MaxBulkExportRowsCeiling = Services.Tools.BulkExportRowCap.Max;
        private const int MinMaxTokens = 1;
        private const int MaxMaxTokens = 200000;

        public static string AdminPassword { get; set; } = "admin";
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

        private static readonly string GlobalConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "OutlookAI",
            "config.xml"
        );

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
            LoadConfigFromPaths(GlobalConfigFilePath, SharedConfigFilePath, UserConfigFilePath);
        }

        public static void LoadConfigFromPaths(string globalConfigPath, string sharedConfigPath, string userConfigPath)
        {
            ResetDefaults();
            LoadFromFile(globalConfigPath, allowServerFields: true, allowApiKey: false);
            LoadFromFile(sharedConfigPath, allowServerFields: false, allowApiKey: false);
            LoadFromFile(userConfigPath, allowServerFields: false, allowApiKey: true);
        }

        public static void LoadConfigFromPaths(string globalConfigPath, string userConfigPath)
        {
            LoadConfigFromPaths(globalConfigPath, sharedConfigPath: null, userConfigPath);
        }

        public static void ResetDefaults()
        {
            AdminPassword = "admin";
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

                var adminPassword = root.Element("AdminPassword");
                if (adminPassword != null && !string.IsNullOrEmpty(adminPassword.Value))
                {
                    AdminPassword = adminPassword.Value;
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
                if (voiceModel != null && !string.IsNullOrWhiteSpace(voiceModel.Value))
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

        public static void SaveConfig()
        {
            var userDoc = new XDocument(
                new XElement("Config",
                    new XElement("AdminPassword", AdminPassword),
                    new XElement("LiteLlmApiKey", LiteLlmApiKey ?? ""),
                    new XElement("ReasoningEffort", ReasoningEffort),
                    new XElement("WriteToolsEnabled", WriteToolsEnabled),
                    new XElement("EnabledWriteTools",
                        string.Join(",", EnabledWriteTools ?? new HashSet<string>()))
                )
            );

            var sharedDoc = new XDocument(
                new XElement("Config",
                    new XElement("AdminPassword", AdminPassword),
                    new XElement("ReasoningEffort", ReasoningEffort),
                    new XElement("WriteToolsEnabled", WriteToolsEnabled),
                    new XElement("EnabledWriteTools",
                        string.Join(",", EnabledWriteTools ?? new HashSet<string>()))
                )
            );

            TrySaveTo(UserConfigFilePath, userDoc);
            TrySaveTo(SharedConfigFilePath, sharedDoc);
        }

        public static string NormalizeBaseUrl(string raw)
        {
            var value = (raw ?? "").Trim();
            while (value.EndsWith("/", StringComparison.Ordinal))
            {
                value = value.Substring(0, value.Length - 1);
            }
            return string.IsNullOrEmpty(value) ? DefaultLiteLlmBaseUrl : value;
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
