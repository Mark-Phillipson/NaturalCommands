using System;
using System.Drawing;
using Xunit;
using NaturalCommands;
using NaturalCommands.Models;
using NaturalCommands.Helpers;
using System.Collections.Generic;

namespace NaturalCommands_NET.Tests
{
    public class VisualIdentifyActionTests
    {
        [Fact]
        public void VisualIdentify_SingleCandidate_AutoClicks()
        {
            // Arrange: ensure visual targeting enabled and set a high threshold.
            // Single-candidate behavior should still auto-click.
            AppSettings.Instance.VisualTargeting.Enabled = true;
            AppSettings.Instance.VisualTargeting.AutoClickConfidenceThreshold = 0.95;

            NaturalLanguageInterpreter.VisualIdentifyCandidatesOverride = (phrase) => new List<VisualTargetCandidate>
            {
                new VisualTargetCandidate
                {
                    Label = "TestTarget",
                    Bounds = new Rectangle(10, 10, 40, 20),
                    Confidence = 0.6,
                    Source = "vision"
                }
            };

            bool clicked = false;
            Point clickedPoint = Point.Empty;
            NaturalLanguageInterpreter.ClickOverride = (pt) => { clicked = true; clickedPoint = pt; };

            var interpreter = new NaturalLanguageInterpreter();

            try
            {
                // Act
                var result = interpreter.ExecuteActionAsync(new VisualIdentifyClickAction("whatever"));

                // Assert: click was invoked and the result acknowledges the click
                Assert.True(clicked, "Expected click override to be invoked for single candidate.");
                Assert.Contains("Clicked visual target", result);
                Assert.Contains("TestTarget", result);

                // Session should be cleared after auto-click
                var session = VisualCandidateSessionStore.GetSession();
                Assert.Null(session);
            }
            finally
            {
                // Cleanup hooks
                NaturalLanguageInterpreter.VisualIdentifyCandidatesOverride = null;
                NaturalLanguageInterpreter.ClickOverride = null;
            }
        }

        [Fact]
        public void VisualIdentify_MultipleVisionCandidates_ShowsNumberedCandidatesWithoutClick()
        {
            // Arrange: multi-candidate with non-UIA top candidate should not auto-click.
            AppSettings.Instance.VisualTargeting.Enabled = true;
            AppSettings.Instance.VisualTargeting.AutoClickConfidenceThreshold = 0.95;

            VisualCandidateSessionStore.Clear();

            NaturalLanguageInterpreter.VisualIdentifyCandidatesOverride = (phrase) => new List<VisualTargetCandidate>
            {
                new VisualTargetCandidate
                {
                    Label = "FirstTarget",
                    Bounds = new Rectangle(30, 40, 50, 20),
                    Confidence = 0.9,
                    Source = "vision"
                },
                new VisualTargetCandidate
                {
                    Label = "SecondTarget",
                    Bounds = new Rectangle(80, 100, 50, 20),
                    Confidence = 0.8,
                    Source = "vision"
                }
            };

            bool clicked = false;
            NaturalLanguageInterpreter.ClickOverride = (pt) => { clicked = true; };

            var interpreter = new NaturalLanguageInterpreter();

            try
            {
                // Act
                var result = interpreter.ExecuteActionAsync(new VisualIdentifyClickAction("thing"));

                // Assert: no click, returns numbered candidate guidance, and session is persisted.
                Assert.False(clicked, "Did not expect click for multi-candidate non-UIA result.");
                Assert.Contains("Found 2 visual candidates", result);
                Assert.Contains("choose 1", result);

                var session = VisualCandidateSessionStore.GetSession();
                Assert.NotNull(session);
                Assert.Equal("thing", session!.Query);
                Assert.Equal(2, session.Candidates.Count);
            }
            finally
            {
                VisualCandidateSessionStore.Clear();
                NaturalLanguageInterpreter.VisualIdentifyCandidatesOverride = null;
                NaturalLanguageInterpreter.ClickOverride = null;
            }
        }

        [Fact]
        public void VisualIdentify_SingleOcrSemanticLabel_AutoClicks()
        {
            AppSettings.Instance.VisualTargeting.Enabled = true;
            AppSettings.Instance.VisualTargeting.AutoClickConfidenceThreshold = 0.97;

            NaturalLanguageInterpreter.VisualIdentifyCandidatesOverride = (phrase) => new List<VisualTargetCandidate>
            {
                new VisualTargetCandidate
                {
                    Label = "Ollama",
                    Bounds = new Rectangle(120, 180, 80, 24),
                    Confidence = 0.95,
                    Source = "ocr"
                }
            };

            bool clicked = false;
            NaturalLanguageInterpreter.ClickOverride = (pt) => { clicked = true; };

            var interpreter = new NaturalLanguageInterpreter();

            try
            {
                var result = interpreter.ExecuteActionAsync(new VisualIdentifyClickAction("llama"));

                Assert.True(clicked, "Expected OCR single semantic label match to auto-click.");
                Assert.Contains("Clicked visual target", result);
                Assert.Contains("Ollama", result);
            }
            finally
            {
                VisualCandidateSessionStore.Clear();
                NaturalLanguageInterpreter.VisualIdentifyCandidatesOverride = null;
                NaturalLanguageInterpreter.ClickOverride = null;
            }
        }
    }
}
