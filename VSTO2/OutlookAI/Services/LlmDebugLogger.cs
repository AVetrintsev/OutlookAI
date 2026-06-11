using System;
using System.IO;
using System.Text;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services
{
    public static class LlmDebugLogger
    {
        private static readonly object Sync = new object();

        public static bool IsEnabled
        {
            get
            {
                return Config.LlmDebugLogEnabled
                    && !string.IsNullOrWhiteSpace(Config.LlmDebugLogPath);
            }
        }

        public static string NewRequestId()
        {
            return DateTimeOffset.UtcNow.ToString("yyyyMMddHHmmssfff")
                + "-" + Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        public static void Write(string requestId, string section, string text)
        {
            if (!IsEnabled)
            {
                return;
            }

            try
            {
                var path = Config.LlmDebugLogPath;
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var sb = new StringBuilder();
                sb.AppendLine();
                sb.AppendLine("================================================================================");
                sb.Append(DateTimeOffset.UtcNow.ToString("o"));
                sb.Append(" request=");
                sb.Append(string.IsNullOrWhiteSpace(requestId) ? "-" : requestId);
                sb.Append(" section=");
                sb.AppendLine(section ?? "");
                sb.AppendLine("--------------------------------------------------------------------------------");
                sb.AppendLine(text ?? "");

                lock (Sync)
                {
                    File.AppendAllText(path, sb.ToString(), Encoding.UTF8);
                }
            }
            catch
            {
                // LLM debug logging must never break the user workflow.
            }
        }

        public static void WriteJson(string requestId, string section, JToken token)
        {
            Write(requestId, section, token == null ? "<null>" : token.ToString(Newtonsoft.Json.Formatting.Indented));
        }
    }
}
