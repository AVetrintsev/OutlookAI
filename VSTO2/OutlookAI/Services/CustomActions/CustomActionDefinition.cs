using System;

namespace OutlookAI.Services.CustomActions
{
    public sealed class CustomActionFile
    {
        public CustomActionDefinition[] Actions { get; set; }
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
    }

    public sealed class CustomActionRunResult
    {
        public string Text { get; set; }
        public string Output { get; set; }
        public string FilePath { get; set; }
        public string DraftId { get; set; }
    }
}
