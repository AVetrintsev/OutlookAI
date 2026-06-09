using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookAI.Services
{
    public sealed class LiteLlmVoiceService : IDisposable
    {
        private readonly LiteLlmCredentialService _credentials;
        private readonly HttpClient _http;
        private readonly bool _ownsHttp;

        public LiteLlmVoiceService(LiteLlmCredentialService credentials)
            : this(credentials, BuildDefaultHttpClient(), ownsHttp: true)
        {
        }

        public LiteLlmVoiceService(LiteLlmCredentialService credentials, HttpClient httpClient, bool ownsHttp = false)
        {
            _credentials = credentials ?? throw new ArgumentNullException(nameof(credentials));
            _http = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _ownsHttp = ownsHttp;
        }

        public static string TranscriptionsEndpoint
            => Config.NormalizeBaseUrl(Config.LiteLlmBaseUrl) + "/audio/transcriptions";

        private static HttpClient BuildDefaultHttpClient()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            return new HttpClient { Timeout = TimeSpan.FromMinutes(2) };
        }

        public async Task<string> TranscribeAsync(Stream pcm, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (pcm == null) throw new ArgumentNullException(nameof(pcm));

            byte[] pcmBytes;
            using (var ms = new MemoryStream())
            {
                await pcm.CopyToAsync(ms, 81920, cancellationToken).ConfigureAwait(false);
                pcmBytes = ms.ToArray();
            }

            var wavBytes = BuildWav(pcmBytes, sampleRate: 16000, bitsPerSample: 16, channels: 1);

            using (var request = new HttpRequestMessage(HttpMethod.Post, TranscriptionsEndpoint))
            using (var form = new MultipartFormDataContent())
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _credentials.GetApiKey());

                form.Add(new StringContent(Config.VoiceModel ?? ""), "model");
                var file = new ByteArrayContent(wavBytes);
                file.Headers.ContentType = new MediaTypeHeaderValue("audio/wav");
                form.Add(file, "file", "speech.wav");
                request.Content = form;

                using (var response = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    var text = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException("LiteLLM transcription error: " + (int)response.StatusCode + " " + text);
                    }

                    var obj = JObject.Parse(text);
                    return ((string)obj["text"] ?? "").Trim();
                }
            }
        }

        private static byte[] BuildWav(byte[] pcm, int sampleRate, short bitsPerSample, short channels)
        {
            pcm = pcm ?? new byte[0];
            var byteRate = sampleRate * channels * bitsPerSample / 8;
            var blockAlign = (short)(channels * bitsPerSample / 8);
            using (var ms = new MemoryStream())
            using (var bw = new BinaryWriter(ms, Encoding.ASCII))
            {
                bw.Write(Encoding.ASCII.GetBytes("RIFF"));
                bw.Write(36 + pcm.Length);
                bw.Write(Encoding.ASCII.GetBytes("WAVE"));
                bw.Write(Encoding.ASCII.GetBytes("fmt "));
                bw.Write(16);
                bw.Write((short)1);
                bw.Write(channels);
                bw.Write(sampleRate);
                bw.Write(byteRate);
                bw.Write(blockAlign);
                bw.Write(bitsPerSample);
                bw.Write(Encoding.ASCII.GetBytes("data"));
                bw.Write(pcm.Length);
                bw.Write(pcm);
                return ms.ToArray();
            }
        }

        public void Dispose()
        {
            if (_ownsHttp)
            {
                _http.Dispose();
            }
        }
    }
}
