using System.Threading;
using System.Threading.Tasks;
using OutlookAI.Services.Chat;

namespace OutlookAI.Services.Assistant
{
    public interface IAssistantEngine
    {
        Task<AssistantEngineResult> RunAsync(
            TurnRequest request,
            ChatEventSink sink,
            CancellationToken cancellationToken);
    }

    public interface ICapabilityIndex
    {
        Task<CapabilitySearchResult[]> SearchAsync(
            string query,
            TurnSnapshot snapshot,
            int topK,
            CancellationToken cancellationToken);
    }

    public interface IEmbeddingProvider
    {
        int Dimension { get; }
        Task<float[]> EmbedAsync(string text, CancellationToken cancellationToken);
    }

    public interface IPlannerClient
    {
        Task<TurnPlan> PlanAsync(
            TurnRequest request,
            TurnSnapshot snapshot,
            CancellationToken cancellationToken);
    }

    public interface IPlanValidator
    {
        ValidatedTurnPlan Validate(
            TurnPlan plan,
            TurnRequest request,
            TurnSnapshot snapshot,
            bool includeWriteTools);
    }

    public interface IContextResolver
    {
        ContextBundle Resolve(
            TurnRequest request,
            TurnSnapshot snapshot,
            ValidatedTurnPlan plan);
    }

    public interface IExecutorClient
    {
        Task<AssistantEngineResult> ExecuteAsync(
            TurnRequest request,
            ContextBundle context,
            ValidatedTurnPlan plan,
            ChatEventSink sink,
            CancellationToken cancellationToken);
    }

    public interface IMemoryManager
    {
        string BuildWorkingContext(TurnRequest request);
        Task<string> CompactAsync(TurnRequest request, CancellationToken cancellationToken);
    }
}
