using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services.Tools
{
    /// <summary>
    /// Tool: outlook_read_message. Fetches one message by id; body always plaintext.
    /// </summary>
    public sealed class OutlookReadMessageTool : IOutlookTool
    {
        public string Name => "outlook_read_message";

        public Task<string> ExecuteAsync(string argsJson, IOutlookSurface surface, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var args = JObject.Parse(argsJson ?? "{}");
            var id = (string)args["message_id"];
            if (string.IsNullOrEmpty(id))
            {
                return Task.FromResult(BuildError("invalid_arguments", "message_id is required"));
            }
            bool includeFullBody = args["include_full_body"]?.Value<bool>() ?? true;

            var detail = surface.ReadMessage(id, includeFullBody);
            if (detail == null)
            {
                return Task.FromResult(BuildError("not_found", "Message " + id + " not found"));
            }

            var json = OutlookJsonProjection.MessageDetail(detail, includeBody: true);
            return Task.FromResult(json.ToString(Newtonsoft.Json.Formatting.None));
        }

        private static string BuildError(string code, string message)
            => new JObject(new JProperty("error",
                new JObject(
                    new JProperty("code", code),
                    new JProperty("message", message))))
               .ToString(Newtonsoft.Json.Formatting.None);
    }
}
