# Talon Interop Architecture

## Overview

NaturalCommands can now route unmatched voice commands to Talon for execution, providing a seamless fallback when native rules don't match. This document explains the current implementation and its limitations.

## Fallback Chain

When a user speaks a command:

1. **Native NaturalCommands Rules** → Check if command matches defined rules
2. **Talon Fallback (Primary for unmatched)** → Fuzzy match + AI semantic resolution
3. **NaturalCommands AI** → General command interpretation (if Talon fails)

## Architecture

### Components

- **[TalonCommandCatalog.cs](Helpers/TalonCommandCatalog.cs)**: Scans Talon user scripts directory and caches available commands
- **[TalonCommandMatcher.cs](Helpers/TalonCommandMatcher.cs)**: Fuzzy token matching + AI semantic resolution of user utterance to Talon commands
- **[TalonCommandRouter.cs](Helpers/TalonCommandRouter.cs)**: Orchestrates catalog and matcher to produce `RunTalonCommandAction`
- **[TalonCommandExecutor.cs](Helpers/TalonCommandExecutor.cs)**: Dispatches matched commands via queue file
- **Python Listener**: (`mystuff/talon_my_stuff/naturalcommands_queue_listener.py`) Polls queue and executes via `actions.mimic()`

### Flow Diagram

```
User Command (spoken or typed)
    ↓
[NaturalLanguageInterpreter] Check native rules
    ↓ (no match)
[TalonCommandRouter] Resolve to Talon command
    ├→ [TalonCommandCatalog] Load available commands
    ├→ [TalonCommandMatcher] Fuzzy + AI match
    ├→ Write to CommandQueueFilePath
    ↓
[CommandQueueFile] (intermediary)
    ↓
[Talon Python Listener] Poll queue, read command
    ↓
[actions.mimic()] Execute command in Talon context
```

## Current Workaround: Queue-Based Dispatch via `actions.mimic()`

### Why This Approach?

**The Problem**: There is no .NET API to directly invoke Talon's voice command system from C#. Talon is a separate Python-based voice control system with its own command registry and context awareness.

**The Solution**: Use a file-based queue as an intermediary:
1. NaturalCommands writes matched command to a queue file
2. Talon's Python listener polls the queue periodically
3. Talon executes the command via `actions.mimic()` in its own runtime context

### Why It's Not Recommended (Limitations)

1. **Indirect Execution**: Command execution is asynchronous and decoupled, making debugging harder
2. **File I/O Latency**: Writing to disk and polling introduces delay (250ms poll interval)
3. **No Direct Error Feedback**: Talon execution errors are not directly reported back to .NET
4. **Mimic() Workaround**: `actions.mimic()` is Talon's internal API for replaying voice commands, not designed for programmatic use
5. **Process Coupling**: Requires both NaturalCommands and Talon to be running and properly configured
6. **Queue File Persistence**: Queue files accumulate and must be periodically cleaned

### Why It's the Only Current Solution

- **No Talon .NET Bindings**: Talon does not provide native libraries for .NET
- **No Public HTTP/Socket API**: Talon's inter-process communication is not standardized
- **Mimic() Is The Closest Match**: `actions.mimic()` is Talon's closest equivalent to "run this voice command"

## Configuration

All settings are in `bin/{Debug,Release}/net10.0-windows/settings.json`:

```json
{
  "Talon": {
    "EnableFallback": true,
    "UserScriptsDirectory": "C:\\Users\\{username}\\AppData\\Roaming\\talon\\user",
    "CatalogCacheFileName": "talon_phrase_catalog.json",
    "MinFallbackMatchScore": 0.7,
    "MinAIConfidence": 0.6,
    "CommandQueueFilePath": "C:\\...\\talon_command_queue.txt",
    "BridgeExecutable": "",
    "BridgeArgumentsTemplate": "{command}"
  }
}
```

## Future Alternatives

### 1. **Talon Socket/HTTP Server**
If Talon adds a public server API, NaturalCommands could send commands directly without file I/O.

### 2. **Talon Python Plugin**
Develop a Talon Python plugin that listens for commands via named pipes or local sockets, eliminating file-based queueing.

### 3. **Native Talon Bindings**
If Talon releases .NET bindings or WASM exports, direct integration would be possible.

### 4. **Custom Talon Bridge**
Build a lightweight Talon application that exposes a stable API for external command execution.

## Testing

### Run a Semantic Talon Match

```powershell
$env:OPENAI_API_KEY = "sk-..."  # Set OpenAI key for semantic matching
dotnet run -- natural "can you show me the talon alphabet"
# Expect: Queue file created with "help alphabet"
```

### Monitor Queue Execution

1. Check queue file: `bin/{Debug,Release}/net10.0-windows/talon_command_queue.txt`
2. Verify Talon listener is running: Check Talon logs
3. Confirm command was executed in Talon

## Known Issues

- **Fuzzy matching can be too broad** with high MinFallbackMatchScore (e.g., "show" matches many commands)
- **AI semantic matching requires OpenAI API** to be available
- **Queue file cleanup** is manual (files are not auto-deleted)
- **No feedback loop** from Talon execution back to NaturalCommands

## See Also

- [talon_my_stuff/naturalcommands_queue_listener.py](../AppData/Roaming/talon/user/mystuff/talon_my_stuff/naturalcommands_queue_listener.py) - Talon listener implementation
- [Talon Documentation](https://talon.wiki/) - Official Talon docs
- [ActionModels.cs](Models/ActionModels.cs) - `RunTalonCommandAction` definition
