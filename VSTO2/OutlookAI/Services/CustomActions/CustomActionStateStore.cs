using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.CustomActions
{
    public sealed class CustomActionStateStore
    {
        public string Path { get; }

        public CustomActionStateStore()
            : this(ManifestPathResolver.LocalAppDataFile("custom-action-state.json"))
        {
        }

        public CustomActionStateStore(string path)
        {
            Path = path ?? throw new ArgumentNullException(nameof(path));
        }

        public DateTimeOffset? GetLastRun(string actionId)
        {
            var map = LoadMap();
            if (map.TryGetValue(actionId ?? "", out var text)
                && DateTimeOffset.TryParse(text, out var dt))
            {
                return dt.ToUniversalTime();
            }
            return null;
        }

        public void SetLastRun(string actionId, DateTimeOffset timestamp)
        {
            var map = LoadMap();
            map[actionId ?? ""] = timestamp.ToUniversalTime().ToString("o");
            var root = new JObject(
                new JProperty("last_run", JObject.FromObject(map)));
            var dir = System.IO.Path.GetDirectoryName(Path);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllText(Path, root.ToString(Formatting.Indented));
        }

        private Dictionary<string, string> LoadMap()
        {
            if (!File.Exists(Path))
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            var root = JObject.Parse(File.ReadAllText(Path));
            var obj = root["last_run"] as JObject;
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (obj == null)
            {
                return map;
            }
            foreach (var prop in obj.Properties())
            {
                map[prop.Name] = (string)prop.Value;
            }
            return map;
        }
    }
}
