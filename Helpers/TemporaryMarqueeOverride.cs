using System;

namespace NaturalCommands.Helpers
{
    public sealed class TemporaryMarqueeOverride
    {
        private long _expiresAtMs;

        public string? Message { get; private set; }

        public void Set(string message, int durationMs, long nowMs)
        {
            Message = message;
            _expiresAtMs = nowMs + Math.Max(0, durationMs);
        }

        public void Clear()
        {
            Message = null;
            _expiresAtMs = 0;
        }

        public bool IsActive(long nowMs)
        {
            if (string.IsNullOrWhiteSpace(Message)) return false;
            return nowMs < _expiresAtMs;
        }

        public string? GetMessage(long nowMs) => IsActive(nowMs) ? Message : null;
    }
}
