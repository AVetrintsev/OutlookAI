using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookAI.Services.Skills
{
    /// <summary>Atomic, revision-checked local knowledge storage shared by all task panes.</summary>
    public sealed class SkillStore
    {
        private static readonly object Gate = new object();
        private const int MaximumFileBytes = 8 * 1024 * 1024;
        public string Path { get; }

        public SkillStore(string path = null)
        {
            Path = path ?? System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "OutlookAI", "skills.json");
        }

        public SkillFile Read()
        {
            lock (Gate)
            {
                if (!File.Exists(Path)) return new SkillFile();
                if (new FileInfo(Path).Length > MaximumFileBytes)
                    throw new InvalidDataException("Файл скиллов превышает 8 МБ.");
                var file = SkillJson.Parse<SkillFile>(File.ReadAllText(Path, Encoding.UTF8));
                ValidateFile(file);
                return file;
            }
        }

        public void AcceptDisclosure()
        {
            lock (Gate)
            {
                var file = Read();
                file.DisclosureAccepted = true;
                Save(file);
            }
        }

        public SkillDefinition Upsert(SkillDefinition skill, int expectedRevision)
        {
            lock (Gate)
            {
                var normalized = SkillValidation.Normalize(skill);
                var file = Read();
                var items = file.Skills.ToList();
                var index = items.FindIndex(item => SameId(item.Id, normalized.Id));
                var currentRevision = index < 0 ? 0 : items[index].Revision;
                if (currentRevision != expectedRevision) throw Conflict();
                normalized.Revision = checked(currentRevision + 1);
                normalized.UpdatedAt = DateTimeOffset.UtcNow;
                if (index < 0) items.Add(normalized); else items[index] = normalized;
                file.Skills = items.ToArray();
                Save(file);
                return normalized;
            }
        }

        public void Delete(string id, int expectedRevision)
        {
            lock (Gate)
            {
                var file = Read();
                var current = file.Skills.FirstOrDefault(skill => SameId(skill.Id, id));
                if (current == null || current.Revision != expectedRevision) throw Conflict();
                file.Skills = file.Skills.Where(skill => !SameId(skill.Id, id)).ToArray();
                Save(file);
            }
        }

        public string Export() => SkillJson.Write(new SkillFile { Skills = Read().Skills });

        public SkillDefinition[] ParseImport(string json)
        {
            if (json == null || Encoding.UTF8.GetByteCount(json) > MaximumFileBytes)
                throw new ArgumentException("Импорт должен содержать JSON размером до 8 МБ.");
            var file = SkillJson.Parse<SkillFile>(json);
            ValidateFile(file);
            return file.Skills.Select(SkillValidation.Normalize).ToArray();
        }

        public void ApplyImport(SkillDefinition[] imported, long expectedFileRevision, IEnumerable<string> replaceIds)
        {
            lock (Gate)
            {
                var incoming = (imported ?? new SkillDefinition[0]).Select(SkillValidation.Normalize).ToArray();
                var file = Read();
                if (file.Revision != expectedFileRevision) throw Conflict();
                var replace = new HashSet<string>(replaceIds ?? new string[0], StringComparer.OrdinalIgnoreCase);
                var items = file.Skills.ToList();
                foreach (var skill in incoming)
                {
                    var index = items.FindIndex(item => SameId(item.Id, skill.Id));
                    if (index >= 0 && !replace.Contains(skill.Id)) continue;
                    skill.Revision = index >= 0 ? checked(items[index].Revision + 1) : 1;
                    skill.UpdatedAt = DateTimeOffset.UtcNow;
                    if (index >= 0) items[index] = skill; else items.Add(skill);
                }
                file.Skills = items.ToArray();
                Save(file);
            }
        }

        private void Save(SkillFile file)
        {
            ValidateFile(file);
            file.Revision = checked(file.Revision + 1);
            var json = SkillJson.Write(file);
            if (Encoding.UTF8.GetByteCount(json) > MaximumFileBytes)
                throw new ArgumentException("Суммарный размер скиллов превышает 8 МБ.");
            var directory = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(Path));
            Directory.CreateDirectory(directory);
            var temp = Path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temp, json, new UTF8Encoding(false));
                if (File.Exists(Path)) File.Replace(temp, Path, Path + ".bak");
                else File.Move(temp, Path);
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }

        private static void ValidateFile(SkillFile file)
        {
            if (file == null || file.SchemaVersion != 1 || file.Skills == null)
                throw new InvalidDataException("Неизвестный формат файла скиллов. Исходный файл не изменён.");
            if (file.Skills.Length > 200) throw new ArgumentException("Допустимо до 200 скиллов.");
            if (file.Skills.Any(skill => skill == null || string.IsNullOrWhiteSpace(skill.Id)))
                throw new ArgumentException("У каждого импортируемого скилла должен быть ID.");
            file.Skills = file.Skills.Select(SkillValidation.Normalize).ToArray();
            if (file.Skills.Select(skill => skill.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count() != file.Skills.Length)
                throw new ArgumentException("Файл содержит повторяющиеся ID скиллов.");
        }

        private static bool SameId(string left, string right) => string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        private static InvalidOperationException Conflict() => new InvalidOperationException(
            "Скиллы изменились в другой панели. Обновите список и повторите действие; ваши изменения не перезаписаны.");
    }
}
