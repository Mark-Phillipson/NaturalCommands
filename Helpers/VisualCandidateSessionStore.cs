using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using NaturalCommands.Models;

namespace NaturalCommands.Helpers
{
    public static class VisualCandidateSessionStore
    {
        private static readonly object _lock = new();
        private static VisualTargetSession? _current;
        private static readonly string _sessionFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "NaturalCommands",
            "visual_candidate_session.json");
        private static readonly TimeSpan _maxSessionAge = TimeSpan.FromMinutes(2);

        private sealed class PersistedCandidate
        {
            public string Label { get; set; } = string.Empty;
            public int X { get; set; }
            public int Y { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
            public double Confidence { get; set; }
            public string Reason { get; set; } = string.Empty;
            public string Source { get; set; } = "uia";
        }

        private sealed class PersistedSession
        {
            public string Query { get; set; } = string.Empty;
            public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;
            public System.Collections.Generic.List<PersistedCandidate> Candidates { get; set; } = new();
        }

        public static void SetSession(VisualTargetSession session)
        {
            lock (_lock)
            {
                _current = session;
                SaveToDisk(session);
            }
        }

        public static VisualTargetSession? GetSession()
        {
            lock (_lock)
            {
                if (_current != null)
                {
                    if (DateTime.UtcNow - _current.CreatedUtc > _maxSessionAge)
                    {
                        _current = null;
                        DeletePersistedSession();
                    }
                    return _current;
                }

                _current = LoadFromDisk();
                return _current;
            }
        }

        public static void Clear()
        {
            lock (_lock)
            {
                _current = null;
                DeletePersistedSession();
            }
        }

        public static bool TryGetCandidate(int oneBasedNumber, out VisualTargetCandidate? candidate)
        {
            lock (_lock)
            {
                candidate = null;
                var session = GetSession();
                if (session == null || oneBasedNumber <= 0)
                {
                    return false;
                }

                candidate = session.Candidates
                    .Skip(oneBasedNumber - 1)
                    .FirstOrDefault();
                return candidate != null;
            }
        }

        private static void SaveToDisk(VisualTargetSession session)
        {
            try
            {
                var dir = Path.GetDirectoryName(_sessionFilePath);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var persisted = new PersistedSession
                {
                    Query = session.Query,
                    CreatedUtc = session.CreatedUtc,
                    Candidates = session.Candidates.Select(c => new PersistedCandidate
                    {
                        Label = c.Label,
                        X = c.Bounds.X,
                        Y = c.Bounds.Y,
                        Width = c.Bounds.Width,
                        Height = c.Bounds.Height,
                        Confidence = c.Confidence,
                        Reason = c.Reason,
                        Source = c.Source
                    }).ToList()
                };

                var json = JsonSerializer.Serialize(persisted);
                File.WriteAllText(_sessionFilePath, json);
            }
            catch
            {
            }
        }

        private static VisualTargetSession? LoadFromDisk()
        {
            try
            {
                if (!File.Exists(_sessionFilePath))
                {
                    return null;
                }

                var json = File.ReadAllText(_sessionFilePath);
                var persisted = JsonSerializer.Deserialize<PersistedSession>(json);
                if (persisted == null)
                {
                    return null;
                }

                if (DateTime.UtcNow - persisted.CreatedUtc > _maxSessionAge)
                {
                    DeletePersistedSession();
                    return null;
                }

                return new VisualTargetSession
                {
                    Query = persisted.Query,
                    CreatedUtc = persisted.CreatedUtc,
                    Candidates = persisted.Candidates.Select(c => new VisualTargetCandidate
                    {
                        Label = c.Label,
                        Bounds = new System.Drawing.Rectangle(c.X, c.Y, c.Width, c.Height),
                        Confidence = c.Confidence,
                        Reason = c.Reason,
                        Source = c.Source
                    }).ToList()
                };
            }
            catch
            {
                return null;
            }
        }

        private static void DeletePersistedSession()
        {
            try
            {
                if (File.Exists(_sessionFilePath))
                {
                    File.Delete(_sessionFilePath);
                }
            }
            catch
            {
            }
        }
    }
}
