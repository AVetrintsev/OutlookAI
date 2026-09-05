namespace OutlookAI.Services.Assistant
{
    public sealed class CapabilityCard
    {
        public string Id { get; set; }
        public string Type { get; set; }
        public string Group { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string[] WhenToUse { get; set; } = new string[0];
        public string[] WhenNotToUse { get; set; } = new string[0];
        public string[] RequiredCapabilities { get; set; } = new string[0];
        public string Risk { get; set; } = "read_only";
        public string[] ToolNames { get; set; } = new string[0];
    }

    public sealed class CapabilitySearchResult
    {
        public CapabilityCard Card { get; set; }
        public double Score { get; set; }
        public string Reason { get; set; } = "";
    }
}
