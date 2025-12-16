using System;

namespace NaturalCommands.Helpers
{
    public sealed class DebounceGate
    {
        private readonly long _debounceMs;
        private long _dueAtMs;
        private bool _pending;

        public DebounceGate(int debounceMs)
        {
            _debounceMs = Math.Max(0, debounceMs);
        }

        public void MarkChange(long nowMs)
        {
            _pending = true;
            _dueAtMs = nowMs + _debounceMs;
        }

        public bool TryConsume(long nowMs)
        {
            if (!_pending) return false;
            if (nowMs < _dueAtMs) return false;
            _pending = false;
            return true;
        }

        public bool Pending => _pending;
        public long DueAtMs => _dueAtMs;
    }
}
