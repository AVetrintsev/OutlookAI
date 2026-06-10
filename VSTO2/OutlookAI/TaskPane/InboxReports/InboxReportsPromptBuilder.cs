namespace OutlookAI.TaskPane.InboxReports
{
    /// <summary>
    /// Builds the system prompt for the Inbox Reports pane. Steers the
    /// model toward concise, structured markdown reports and tells it
    /// when to prefer the bulk-read and aggregate tools.
    /// </summary>
    public sealed class InboxReportsPromptBuilder
    {
        public string Build()
        {
            return OutlookAI.Services.PromptCatalog.Default.Get("inbox_reports");
        }
    }
}
