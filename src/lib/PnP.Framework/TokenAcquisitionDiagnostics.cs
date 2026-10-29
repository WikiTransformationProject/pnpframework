using Microsoft.Identity.Client;

namespace PnP.Framework
{
    // HEU: written by LLM, 2026-08-10, 2026-10-29
    public enum TokenAcquisitionFlow
    {
        Silent,
        Interactive,
        DeviceCode,
        UsernamePassword
    }

    // HEU: written by LLM, 2026-08-10, 2026-10-29
    public sealed class TokenAcquisitionDiagnostics
    {
        public TokenAcquisitionDiagnostics(
            TokenAcquisitionFlow flow,
            bool succeeded,
            int cachedAccountCount,
            string accountUsername,
            TokenSource? tokenSource,
            string failureType,
            string errorCode,
            TokenAcquisitionFlow? fallbackFlow = null)
        {
            Flow = flow;
            Succeeded = succeeded;
            CachedAccountCount = cachedAccountCount;
            AccountUsername = accountUsername;
            TokenSource = tokenSource;
            FailureType = failureType;
            ErrorCode = errorCode;
            FallbackFlow = fallbackFlow;
        }

        public TokenAcquisitionFlow Flow { get; }
        public bool Succeeded { get; }
        public int CachedAccountCount { get; }
        public string AccountUsername { get; }
        public TokenSource? TokenSource { get; }
        public string FailureType { get; }
        public string ErrorCode { get; }
        public TokenAcquisitionFlow? FallbackFlow { get; }
    }
}
