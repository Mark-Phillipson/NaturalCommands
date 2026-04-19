using System;
using System.Drawing;

namespace NaturalCommands.Models
{
    public class VisualTargetCandidate
    {
        public string Label { get; set; } = string.Empty;
        public Rectangle Bounds { get; set; } = Rectangle.Empty;
        public double Confidence { get; set; }
        public string Reason { get; set; } = string.Empty;
        public string Source { get; set; } = "uia";
        public string CandidateId { get; set; } = Guid.NewGuid().ToString();

        public Point Center => new Point(Bounds.Left + (Bounds.Width / 2), Bounds.Top + (Bounds.Height / 2));
    }

    public class VisualTargetSession
    {
        public string Query { get; set; } = string.Empty;
        public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;
        public System.Collections.Generic.List<VisualTargetCandidate> Candidates { get; set; } = new();
    }
}
