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
                Disabled = Disabled
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
    }
}
