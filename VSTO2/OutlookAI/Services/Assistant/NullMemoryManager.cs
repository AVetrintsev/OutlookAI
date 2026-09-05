using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Services.Assistant
{
    public sealed class NullMemoryManager : IMemoryManager
    {
        public string BuildWorkingContext(TurnRequest request)
        {
            return "";
        }

        public Task<string> CompactAsync(TurnRequest request, CancellationToken cancellationToken)
        {
            return Task.FromResult("");
        }
    }
}
