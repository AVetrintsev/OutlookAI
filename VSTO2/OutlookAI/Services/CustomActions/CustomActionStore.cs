using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace OutlookAI.Services.CustomActions
{
    public sealed class CustomActionStore
    {
        public string Path { get; }

        public CustomActionStore()
            : this(ManifestPathResolver.AppDataFile("custom-actions.json"))
        {
        }

        public CustomActionStore(string path)
        {
            Path = path ?? throw new ArgumentNullException(nameof(path));
        }

        public IReadOnlyList<CustomActionDefinition> Load()
        {
            EnsureFileExists();
            var file = JsonConvert.DeserializeObject<CustomActionFile>(File.ReadAllText(Path))
                ?? new CustomActionFile();
            return (file.Actions ?? new CustomActionDefinition[0])
                .Where(IsValid)
                .ToList();
        }

        public void Save(IEnumerable<CustomActionDefinition> actions)
        {
            var file = new CustomActionFile
            {
                Actions = (actions ?? Enumerable.Empty<CustomActionDefinition>())
                    .Where(IsValid)
                    .ToArray()
            };
            File.WriteAllText(Path, JsonConvert.SerializeObject(file, Formatting.Indented));
        }

        public void Upsert(CustomActionDefinition action)
        {
            if (!IsValid(action))
            {
                throw new ArgumentException("Custom action id, title, prompt and context are required.", nameof(action));
            }

            var actions = Load().ToList();
            var idx = actions.FindIndex(a => string.Equals(a.Id, action.Id, StringComparison.OrdinalIgnoreCase));
            if (idx >= 0)
            {
                actions[idx] = action;
            }
            else
            {
                actions.Add(action);
            }
            Save(actions);
        }

        private void EnsureFileExists()
        {
            var dir = System.IO.Path.GetDirectoryName(Path);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }
            if (!File.Exists(Path))
            {
                Save(new[]
                {
                    new CustomActionDefinition
                    {
                        Id = "summarize_current_thread",
                        Title = "Сводка переписки",
                        Prompt = "Сделай краткую сводку переписки. Выдели суть, что требуется от меня, сроки и риски.",
                        Context = new CustomActionContext
                        {
                            Source = "current_selection",
                            MessageScope = "thread",
                            FolderScope = "current_folder",
                            ReadFilter = "all",
                            TimeRange = "today",
                            IncludeFullBodies = true,
                            MaxItems = 20
                        },
                        Output = "chat",
                        AllowTools = false
                    }
                });
            }
        }

        private static bool IsValid(CustomActionDefinition action)
        {
            return action != null
                && !string.IsNullOrWhiteSpace(action.Id)
                && !string.IsNullOrWhiteSpace(action.Title)
                && !string.IsNullOrWhiteSpace(action.Prompt)
                && action.Context != null;
        }
    }
}
