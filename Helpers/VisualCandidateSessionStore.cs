using System;
using System.Linq;
using NaturalCommands.Models;

namespace NaturalCommands.Helpers
{
    public static class VisualCandidateSessionStore
    {
        private static readonly object _lock = new();
        private static VisualTargetSession? _current;

        public static void SetSession(VisualTargetSession session)
        {
            lock (_lock)
            {
                _current = session;
            }
        }

        public static VisualTargetSession? GetSession()
        {
            lock (_lock)
            {
                return _current;
            }
        }

        public static void Clear()
        {
            lock (_lock)
            {
                _current = null;
            }
        }

        public static bool TryGetCandidate(int oneBasedNumber, out VisualTargetCandidate? candidate)
        {
            lock (_lock)
            {
                candidate = null;
                if (_current == null || oneBasedNumber <= 0)
                {
                    return false;
                }

                candidate = _current.Candidates
                    .Skip(oneBasedNumber - 1)
                    .FirstOrDefault();
                return candidate != null;
            }
        }
    }
}
