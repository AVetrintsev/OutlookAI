namespace OutlookAI.Services.Chat
{
    /// <summary>
    /// Why <see cref="OutlookAI.Services.LiteLlmChatService"/> stopped a turn.
    /// </summary>
    public enum StopReason
    {
        Completed,
        Cancelled,
        MaxRoundsReached,
        Error,
    }
}
