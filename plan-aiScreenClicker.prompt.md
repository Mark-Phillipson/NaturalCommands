## Plan: AI Screen Clicker with Budget Fallback (DRAFT)

Build a new voice-driven “visual target click” flow that captures all monitors, sends image + phrase to OpenAI vision for candidate bounding boxes, auto-clicks when confidence is high, and otherwise shows a numbered overlay for disambiguation. The design targets sub-1s perceived latency by using a tiered strategy: fast local pre-filtering and strict screenshot sizing, then one lightweight vision call, escalating only when needed. Because your budget is under $10/month, the plan includes request throttling, token/call caps, and a local fallback chain (UI Automation first, OCR second) to keep cloud usage rare. The implementation follows repo conventions by adding new helper classes instead of growing interpreter complexity.

**Steps**
1. Define command/action contracts for visual targeting in `ActionModels.cs` and `CommandDefinitions.cs`, including actions for `VisualIdentifyClick`, `VisualShowCandidates`, and `VisualChooseCandidate`.
2. Extend intent parsing and follow-up selection routing in `NaturalLanguageInterpreter.cs` around existing quick-click patterns, adding support for utterances like “identify ten of spades” and “choose 2”.
3. Add a screenshot service in `Helpers/` (new file, e.g., `ScreenCaptureService`) that captures all monitors, normalizes scale/DPI, and produces a compressed image budget profile (fast/medium quality tiers).
4. Add a vision grounding service in `Helpers/AIInterpreter.cs` or a new helper (preferred) that calls OpenAI multimodal with structured JSON output for candidates (`label`, `bbox`, `confidence`, `reason`) and returns screen-space coordinates.
5. Implement candidate session state (new helper/model under `Helpers/` and/or `Models/`) to store the latest candidate list and map numeric selection to coordinates for the immediate next voice command.
6. Create a numbered overlay form by reusing patterns from `UIElementOverlayForm.cs` and `QuickClickOverlayForm.cs`, rendering 1..N markers at candidate centers with timeout and cancel behavior.
7. Implement execution policy in interpreter/action executor: auto-click only if single candidate and confidence ≥ threshold; otherwise show overlay and wait for `choose N`; click via existing mouse helpers.
8. Add budget/latency guardrails in settings model/UI: new options in `Models/AppSettings.cs` and `SettingsForm.cs` for max visual calls/day, model tier, confidence threshold, and fallback mode; also fix model persistence inconsistency noted in settings save path.
9. Implement fallback chain for cost/latency resilience: first UI Automation discovery via `Helpers/UIAutomationHelper.cs`, then OCR (new local OCR helper) when vision call is skipped/fails/exceeds budget.
10. Document behavior and cost controls in `README.md` and/or `AUTO_CLICK_USAGE.md`, including examples and expected latency/cost tradeoffs.

**Verification**
- Build: run `dotnet build NaturalCommands.csproj`.
- Functional checks:
  - Voice: “identify ten of spades” with multiple matches shows numbered overlay.
  - Voice: “choose 2” clicks second candidate from active session.
  - Single high-confidence match auto-clicks without overlay.
  - Budget cap reached forces fallback path without OpenAI call.
- Performance checks:
  - Measure end-to-end p50/p95 latency against <1s target on representative screens.
  - Track cloud call count/day and estimated token spend with configured caps.
- Reliability checks:
  - Multi-monitor coordinate mapping accuracy.
  - DPI scaling correctness.
  - Overlay cleanup on timeout/cancel/app switch.

**Decisions**
- Capture scope: all monitors.
- Interaction: auto-click on high confidence, disambiguate otherwise.
- Latency target: under 1 second.
- Cost target: under $10/month via caps + fallback chain.
- Architecture choice: new helpers/forms and minimal interpreter branching, consistent with repo guidance in `.github/copilot.instructions.md`.