using System;
using System.Linq;
using System.Net.Mail;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace OutlookAI.Services.Skills
{
    public sealed class SkillDefinition
    {
        public string Id { get; set; }
        public int Revision { get; set; }
        public string Category { get; set; } = "rules";
        public string Name { get; set; }
        public string Description { get; set; }
        public string[] Triggers { get; set; } = new string[0];
        public string Instructions { get; set; }
        public SkillPerson[] People { get; set; } = new SkillPerson[0];
        public string Project { get; set; }
        public string Client { get; set; }
        public DateTime? ValidFrom { get; set; }
        public DateTime? ValidUntil { get; set; }
        public bool Enabled { get; set; } = true;
        public SkillSource[] Sources { get; set; } = new SkillSource[0];
        public DateTimeOffset UpdatedAt { get; set; }

        public bool IsActive(DateTime today) => Enabled
            && (!ValidFrom.HasValue || ValidFrom.Value.Date <= today.Date)
            && (!ValidUntil.HasValue || ValidUntil.Value.Date >= today.Date);

        public SkillDefinition Clone() => SkillJson.Parse<SkillDefinition>(SkillJson.Write(this));
    }

    public sealed class SkillPerson
    {
        public string Email { get; set; }
        public string Name { get; set; }
        public string[] Aliases { get; set; } = new string[0];
        public string Role { get; set; }
        public string[] Responsibilities { get; set; } = new string[0];
    }

    public sealed class SkillSource
    {
        public string MessageId { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public DateTimeOffset? Date { get; set; }
        public string Note { get; set; }
    }

    public sealed class SkillFile
    {
        public int SchemaVersion { get; set; } = 1;
        public long Revision { get; set; }
        public bool DisclosureAccepted { get; set; }
        public SkillDefinition[] Skills { get; set; } = new SkillDefinition[0];
    }

    public sealed class SkillDraft
    {
        public SkillDefinition Skill { get; set; }
        public string Reason { get; set; }
        public string[] Uncertainties { get; set; } = new string[0];
    }

    public static class SkillJson
    {
        private static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            ContractResolver = new DefaultContractResolver { NamingStrategy = new SnakeCaseNamingStrategy() },
            MaxDepth = 24,
            TypeNameHandling = TypeNameHandling.None
        };

        public static string Write(object value) => JsonConvert.SerializeObject(value, Formatting.Indented, Settings);
        public static T Parse<T>(string json) => JsonConvert.DeserializeObject<T>(json, Settings);
    }

    public static class SkillValidation
    {
        public static SkillDefinition Normalize(SkillDefinition value)
        {
            if (value == null) throw new ArgumentException("Скилл не задан.");
            var skill = value.Clone();
            skill.Id = Text(skill.Id, 100);
            if (skill.Id.Length == 0) skill.Id = Guid.NewGuid().ToString("N");
            if (!skill.Id.All(ch => char.IsLetterOrDigit(ch) || ch == '-' || ch == '_'))
                throw new ArgumentException("ID скилла содержит недопустимые символы.");
            skill.Name = Required(skill.Name, 100, "Название");
            skill.Description = Required(skill.Description, 600, "Описание");
            skill.Instructions = Required(skill.Instructions, 32000, "Инструкции");
            skill.Category = Text(skill.Category, 30).ToLowerInvariant();
            if (!new[] { "people", "rules", "communication" }.Contains(skill.Category))
                throw new ArgumentException("Неизвестный раздел скилла.");
            skill.Triggers = Strings(skill.Triggers, 40, 200);
            skill.Project = Text(skill.Project, 200);
            skill.Client = Text(skill.Client, 200);
            if (skill.ValidFrom.HasValue && skill.ValidUntil.HasValue && skill.ValidFrom.Value.Date > skill.ValidUntil.Value.Date)
                throw new ArgumentException("Начало действия скилла позже окончания.");
            skill.People = skill.People ?? new SkillPerson[0];
            if (skill.People.Length > 100 || skill.People.Any(person => person == null))
                throw new ArgumentException("Допустимо до 100 сотрудников в скилле.");
            foreach (var person in skill.People)
            {
                person.Email = Required(person.Email, 254, "Email сотрудника").ToLowerInvariant();
                try
                {
                    if (!person.Email.Contains("@") || new MailAddress(person.Email).Address != person.Email)
                        throw new FormatException();
                }
                catch (FormatException) { throw new ArgumentException("Проверьте email сотрудника: " + person.Email); }
                person.Name = Required(person.Name, 200, "Имя сотрудника");
                person.Role = Text(person.Role, 300);
                person.Aliases = Strings(person.Aliases, 20, 200);
                person.Responsibilities = Strings(person.Responsibilities, 40, 1000);
            }
            if (skill.People.Select(p => p.Email).Distinct(StringComparer.OrdinalIgnoreCase).Count() != skill.People.Length)
                throw new ArgumentException("Email сотрудников внутри скилла должны быть уникальны.");
            skill.Sources = skill.Sources ?? new SkillSource[0];
            if (skill.Sources.Length > 30 || skill.Sources.Any(source => source == null))
                throw new ArgumentException("Допустимо до 30 источников скилла.");
            foreach (var source in skill.Sources)
            {
                source.MessageId = Text(source.MessageId, 200);
                source.Subject = Text(source.Subject, 400);
                source.From = Text(source.From, 300);
                source.Note = Text(source.Note, 1000);
            }
            return skill;
        }

        private static string Text(string text, int max)
        {
            text = (text ?? "").Trim();
            if (text.Length > max) throw new ArgumentException("Поле превышает допустимую длину: " + max + " символов.");
            return text;
        }

        private static string Required(string text, int max, string label)
        {
            text = Text(text, max);
            if (text.Length == 0) throw new ArgumentException("Заполните поле «" + label + "».");
            return text;
        }

        private static string[] Strings(string[] values, int maxCount, int maxLength)
        {
            values = values ?? new string[0];
            if (values.Length > maxCount) throw new ArgumentException("Слишком много значений в поле.");
            return values.Select(value => Text(value, maxLength)).Where(value => value.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        }
    }
}
