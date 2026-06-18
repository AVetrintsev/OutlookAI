using System;
using System.Linq;

namespace OutlookAI.Services.CustomActions
{
    public sealed class CustomActionFile
    {
        public int SchemaVersion { get; set; } = 2;
        public CustomActionGroup[] Groups { get; set; }

        // Schema v1 compatibility. The store migrates this array on load.
        public CustomActionDefinition[] Actions { get; set; }
    }

    public sealed class CustomActionGroup
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public int Order { get; set; }
        public CustomActionDefinition[] Actions { get; set; }

        public CustomActionGroup Clone()
        {
            return new CustomActionGroup
            {
                Id = Id,
                Title = Title,
                Order = Order,
                Actions = (Actions ?? new CustomActionDefinition[0])
                    .Select(action => action.Clone())
                    .ToArray()
            };
        }
    }

    public sealed class CustomActionDefinition
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Prompt { get; set; }
        public CustomActionContext Context { get; set; }
        public string Output { get; set; }
        public bool AllowTools { get; set; }
        public string[] AllowedTools { get; set; }
        public bool Disabled { get; set; }
        public string ApplicabilityItemType { get; set; }
        public string ApplicabilityDirection { get; set; }

        public CustomActionDefinition Clone()
        {
            return new CustomActionDefinition
            {
                Id = Id,
                Title = Title,
                Description = Description,
                Prompt = Prompt,
                Context = Context?.Clone(),
                Output = Output,
                AllowTools = AllowTools,
                AllowedTools = (AllowedTools ?? new string[0]).ToArray(),
                Disabled = Disabled,
                ApplicabilityItemType = ApplicabilityItemType,
                ApplicabilityDirection = ApplicabilityDirection
            };
        }
    }

    public sealed class CustomActionContext
    {
        public string Source { get; set; }
        public string MessageScope { get; set; }
        public string FolderScope { get; set; }
        public string ReadFilter { get; set; }
        public string TimeRange { get; set; }
        public DateTimeOffset? ManualFrom { get; set; }
        public DateTimeOffset? ManualTo { get; set; }
        public bool IncludeFullBodies { get; set; }
        public bool IncludeAttachments { get; set; }
        public int MaxItems { get; set; }

        public CustomActionContext Clone()
        {
            return (CustomActionContext)MemberwiseClone();
        }
    }

    public sealed class CustomActionRunResult
    {
        public string Text { get; set; }
        public string Output { get; set; }
        public string FilePath { get; set; }
        public string DraftId { get; set; }
        public string DraftLocation { get; set; }
        public string DraftDisplayName { get; set; }
    }

    public sealed class CustomActionApplicabilityContext
    {
        public string ItemType { get; set; }
        public string Direction { get; set; }
    }

    public static class CustomActionApplicability
    {
        public const string All = "all";
        public const string Mail = "mail";
        public const string Meeting = "meeting";
        public const string Incoming = "incoming";
        public const string Outgoing = "outgoing";
        public const string Unknown = "unknown";

        public static bool IsApplicable(CustomActionDefinition action, CustomActionApplicabilityContext context)
        {
            if (action == null || action.Disabled) return false;
            var actionItemType = NormalizeApplicability(action.ApplicabilityItemType, All, Mail, Meeting);
            var actionDirection = NormalizeApplicability(action.ApplicabilityDirection, All, Incoming, Outgoing);
            var itemType = NormalizeContext(context?.ItemType);
            var direction = NormalizeContext(context?.Direction);
            return Matches(actionItemType, itemType) && Matches(actionDirection, direction);
        }

        public static CustomActionFile FilterCatalog(CustomActionFile catalog, CustomActionApplicabilityContext context)
        {
            return new CustomActionFile
            {
                SchemaVersion = catalog?.SchemaVersion ?? CustomActionStore.CurrentSchemaVersion,
                Groups = (catalog?.Groups ?? new CustomActionGroup[0])
                    .Select(group =>
                    {
                        var copy = group.Clone();
                        copy.Actions = (copy.Actions ?? new CustomActionDefinition[0])
                            .Where(action => IsApplicable(action, context))
                            .ToArray();
                        return copy;
                    })
                    .Where(group => group.Actions.Length > 0)
                    .ToArray()
            };
        }

        public static string NormalizeItemType(string value)
        {
            return NormalizeApplicability(value, All, Mail, Meeting);
        }

        public static string NormalizeDirection(string value)
        {
            return NormalizeApplicability(value, All, Incoming, Outgoing);
        }

        private static bool Matches(string actionValue, string contextValue)
        {
            if (actionValue == All) return true;
            if (contextValue == Unknown) return false;
            return string.Equals(actionValue, contextValue, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeContext(string value)
        {
            value = (value ?? "").Trim().ToLowerInvariant();
            return string.IsNullOrWhiteSpace(value) ? Unknown : value;
        }

        private static string NormalizeApplicability(string value, string fallback, params string[] allowed)
        {
            value = (value ?? "").Trim().ToLowerInvariant();
            if (value == All) return All;
            return allowed.Any(item => string.Equals(item, value, StringComparison.OrdinalIgnoreCase))
                ? value
                : fallback;
        }
    }
}
