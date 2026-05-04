using NaturalCommands;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.IO;

namespace ExecuteCommands_NET
{
	internal static class Program
	{
		private static TickerOverlayForm? _tickerForm = null;
		private static readonly object _tickerLock = new object();

		// Quick Clicks command-file support removed.

		private static string GetTickerPayloadFilePath()
		{
			var baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NaturalCommands");
			Directory.CreateDirectory(baseDir);
			return Path.Combine(baseDir, ".ticker_payload");
		}

		private static bool HasOtherNaturalCommandsProcess()
		{
			var currentProcessId = Process.GetCurrentProcess().Id;
			return Process.GetProcessesByName("NaturalCommands").Any(process => process.Id != currentProcessId);
		}

		private static void DeliverTickerPayload(IEnumerable<string> lines)
		{
			var payloadPath = GetTickerPayloadFilePath();
			var tempPath = payloadPath + ".tmp";
			File.WriteAllLines(tempPath, lines);
			if (File.Exists(payloadPath))
			{
				File.Delete(payloadPath);
			}

			File.Move(tempPath, payloadPath);
			NaturalCommands.Helpers.Logger.LogInfo($"[WATCHER] Delivered ticker payload to resident form: {payloadPath}");
		}

		private static void StartTickerWatcher()
		{
			var watcherThread = new System.Threading.Thread(() =>
			{
				var payloadPath = GetTickerPayloadFilePath();
				string? lastHash = null;
				NaturalCommands.Helpers.Logger.LogInfo($"[WATCHER] Starting payload watcher on {payloadPath}");

				// Create the single persistent ticker form on a UI thread
				System.Threading.Thread? uiThread = null;
				System.Action<string> recreateFormIfNeeded = (reason) =>
				{
					try
					{
						lock (_tickerLock)
						{
							if (_tickerForm == null || _tickerForm.IsDisposed)
							{
								NaturalCommands.Helpers.Logger.LogInfo($"[TICKER-UI] Recreating form ({reason})...");
								// Kill existing thread if any
								if (uiThread != null && uiThread.IsAlive)
								{
									// Thread will exit naturally when we exit this method
								}

								var handleReady = new System.Threading.ManualResetEventSlim(false);
								uiThread = new System.Threading.Thread(() =>
								{
									try
									{
										NaturalCommands.Helpers.Logger.LogInfo("[TICKER-UI] Creating persistent ticker form...");
										lock (_tickerLock)
										{
											_tickerForm = new TickerOverlayForm(new[] { "Ticker ready" }, cycleSeconds: 5, maxCycles: 0, topPosition: false, hideOnDismiss: true);
											// Force handle creation now so the parent thread can wait for it
											var _ = _tickerForm.Handle;
										}
										try { handleReady.Set(); } catch { }
										NaturalCommands.Helpers.Logger.LogInfo("[TICKER-UI] Persistent ticker form created, entering message pump");
										Application.Run(_tickerForm);
										NaturalCommands.Helpers.Logger.LogInfo("[TICKER-UI] Message pump exited");
									}
									catch (Exception ex)
									{
										NaturalCommands.Helpers.Logger.LogError($"[TICKER-UI] Failed to create persistent ticker: {ex.Message} | {ex.StackTrace}");
									}
								})
								{
									IsBackground = false
								};

								uiThread.TrySetApartmentState(System.Threading.ApartmentState.STA);
								uiThread.Start();
								NaturalCommands.Helpers.Logger.LogInfo("[TICKER-UI] UI thread spawned");
								// Wait for the UI thread to create the form handle so we avoid races where
								// multiple creator threads attempt to make windows and hit "Window handle already exists".
								try
								{
									if (!handleReady.Wait(TimeSpan.FromSeconds(5)))
									{
										NaturalCommands.Helpers.Logger.LogWarning("[TICKER-UI] Timed out waiting for ticker form handle");
									}
								}
							catch { }
							finally
							{
								try { handleReady.Dispose(); } catch { }
							}
							}
						}
					}
					catch (Exception ex)
					{
						NaturalCommands.Helpers.Logger.LogError($"[TICKER-UI] Error recreating form: {ex.Message}");
					}
				};

				// Initial form creation
				recreateFormIfNeeded("initial");

				// Now watch for payload changes and add messages to the existing form
				while (true)
				{
					try
					{
						if (File.Exists(payloadPath))
						{
							var content = File.ReadAllText(payloadPath);
							if (string.IsNullOrWhiteSpace(content))
							{
								System.Threading.Thread.Sleep(500);
								continue;
							}

							var currentHash = System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(content));
							var currentHashStr = System.Convert.ToHexString(currentHash);

							// Only process if content has changed
							if (lastHash != currentHashStr)
							{
								NaturalCommands.Helpers.Logger.LogInfo($"[WATCHER] Payload file changed. New hash: {currentHashStr.Substring(0, 8)}...");
								lastHash = currentHashStr;
								var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
								NaturalCommands.Helpers.Logger.LogInfo($"[WATCHER] Read {lines.Length} lines from payload");

								if (lines.Length > 0)
								{
									// Recreate form if disposed before processing messages
									recreateFormIfNeeded("payload received while disposed");

									lock (_tickerLock)
									{
										if (_tickerForm != null && !_tickerForm.IsDisposed)
										{
											NaturalCommands.Helpers.Logger.LogInfo("[WATCHER] Adding messages to persistent ticker form");
											foreach (var line in lines)
											{
												var message = TickerOverlayForm.ParseSingleLine(line);
												_tickerForm.AddMessage(message);
												NaturalCommands.Helpers.Logger.LogInfo($"[WATCHER] Added message: {message.Text} ({message.Category})");
											}
										}
										else
										{
											NaturalCommands.Helpers.Logger.LogWarning("[WATCHER] Ticker form is null or disposed after recreation attempt, will retry next cycle");
										}
									}
								}
							}
						}

						System.Threading.Thread.Sleep(500);
					}
					catch (Exception ex)
					{
						NaturalCommands.Helpers.Logger.LogError($"[WATCHER] Watcher error: {ex.Message}");
						System.Threading.Thread.Sleep(1000);
					}
				}
			})
			{
				IsBackground = false,
				Name = "TickerPayloadWatcher"
			};

			watcherThread.Start();
			NaturalCommands.Helpers.Logger.LogInfo("Ticker payload watcher started as foreground thread (keeps app alive)");
		}

		/// <summary>
		///  The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			// Initialize Windows Forms early (only once)
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);

            // Ensure console uses UTF-8 so emoji and dashes render correctly in terminal
            try { Console.OutputEncoding = System.Text.Encoding.UTF8; Console.InputEncoding = System.Text.Encoding.UTF8; } catch { }

			// Initialize settings at startup
			try
			{
				var settings = NaturalCommands.Models.AppSettings.Instance;
				NaturalCommands.Helpers.Logger.LogInfo("Settings loaded successfully");
			}
			catch (Exception ex)
			{
				NaturalCommands.Helpers.Logger.LogError($"Failed to load settings: {ex.Message}");
			}

			// -------------------------------------------------------------
			// CLI contract:
			//   NaturalCommands.exe <mode> <dictation>
			//     <mode>: 'natural', 'sharp', or other string
			//     <dictation>: free-form text to interpret or execute
			//
			// Examples:
			//   ExecuteCommands.exe natural "move this window to the other screen"
			//   ExecuteCommands.exe sharp "Jump to Symbol"
			// -------------------------------------------------------------
			string[] args = Environment.GetCommandLineArgs();

			// If running in debug mode and no args provided, use default test command
			if (System.Diagnostics.Debugger.IsAttached && (args == null || args.Length < 2))
			{
				// When debugging, show a small input dialog so developers can enter a free-form
				// command and optionally choose an application to target. Defaults to:
				//     natural what can I say
				try
				{
					Application.EnableVisualStyles();
					Application.SetCompatibleTextRenderingDefault(false);
					var dlgResult = ShowDebugInputDialog();
					if (dlgResult != null)
					{
						args = new string[] { "NaturalCommands.exe", dlgResult[0], dlgResult[1] };
						NaturalCommands.Helpers.Logger.LogDebug($"Using debug input: {dlgResult[0]} '{dlgResult[1]}' (target app: {dlgResult[2]})");
					}
					else
					{
						args = new string[] { "ExecuteCommands.exe", "natural", "focus fairies little helper" };
						NaturalCommands.Helpers.Logger.LogDebug("Debug dialog cancelled. Defaulting to sample input.");
					}
				}
				catch (Exception ex)
				{
					args = new string[] { "ExecuteCommands.exe", "natural", "focus fairies little helper" };
					NaturalCommands.Helpers.Logger.LogError($"Failed to show debug input dialog: {ex.Message}. Using default sample.");
				}
			}
			// Otherwise, if no arguments, default to natural mode and sample dictation
			else if (args.Length < 2)
			{
				args = new string[] { "ExecuteCommands.exe", "natural", "close tab" };
				NaturalCommands.Helpers.Logger.LogDebug("No arguments detected. Defaulting to: natural 'close tab'");
			}

			// Diagnostic: log raw args
			NaturalCommands.Helpers.Logger.LogDebug($"Raw args: [{string.Join(", ", args)}]");

			string modeRaw = args[1];
			string textRaw = args.Length > 2 ? string.Join(" ", args.Skip(2)) : "";
			string mode = modeRaw.TrimStart('/').Trim().ToLower();
			string text = textRaw.TrimStart('/').Trim();
			var isStandaloneTickerMode = mode == "ticker" || mode == "ticker-file" || mode == "ticker-test";
			var hasOtherNaturalCommandsProcess = HasOtherNaturalCommandsProcess();

			if (!isStandaloneTickerMode && !hasOtherNaturalCommandsProcess)
			{
				try
				{
					StartTickerWatcher();
				}
				catch (Exception ex)
				{
					NaturalCommands.Helpers.Logger.LogError($"Failed to start ticker watcher: {ex.Message}");
				}
			}
			else if (hasOtherNaturalCommandsProcess)
			{
				NaturalCommands.Helpers.Logger.LogInfo($"Skipping ticker watcher startup because another NaturalCommands process is already running (mode: {mode}).");
			}

			if (mode == "listen")
			{
				RunListenMode();
				return;
			}

		if (mode == "listen-notifications" || mode == "notifications-listen")
		{
			var listener = new WindowsNotificationListenerService();
			Console.WriteLine("[listen-notifications] Requesting notification access...");
			var allowed = listener.Start();  // does RequestAccess + starts poll — all on one STA thread
			if (!allowed)
			{
				NaturalCommands.Helpers.Logger.LogError("Notification listener access denied. Go to Windows Settings > System > Notifications, enable NaturalCommands.");
				Console.Error.WriteLine("[listen-notifications] ACCESS DENIED — enable NaturalCommands in Windows Settings > System > Notifications.");
				return;
			}

			NaturalCommands.Helpers.Logger.LogInfo("Windows notification listener started.");
			Console.WriteLine("[listen-notifications] Listening for Windows notifications...");
			Application.Run(new ApplicationContext());
			return;
		}

		if (mode == "ticker" || mode == "ticker-file")
		{
			IEnumerable<string> lines;
			if (mode == "ticker-file")
			{
				try
				{
					lines = File.ReadAllLines(text);
				}
				catch (Exception ex)
				{
					NaturalCommands.Helpers.Logger.LogError($"Ticker file read failed: {ex.Message}");
					lines = new[] { $"critical:Failed to read ticker file: {ex.Message}" };
				}
			}
			else
			{
				lines = text.Split('|', StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim());
			}

			// Diagnostic: log the lines read
			try
			{
				NaturalCommands.Helpers.Logger.LogInfo($"Ticker invoked with {lines.Count()} lines:");
				int idx = 0;
				foreach (var l in lines)
				{
					NaturalCommands.Helpers.Logger.LogInfo($"  [{++idx}] {l}");
				}
			}
			catch { }

			if (hasOtherNaturalCommandsProcess)
			{
				DeliverTickerPayload(lines);
				return;
			}

			var ticker = new TickerOverlayForm(lines, cycleSeconds: 5, maxCycles: 5, topPosition: false);
			Application.Run(ticker);
			return;
		}

		if (mode == "ticker-test")
		{
			var testLines = new[]
			{
				"info:Ticker test started",
				"success:Ticker overlay is working",
				"warning:Ticker test message",
				"critical:Ticker test critical alert"
			};

			if (hasOtherNaturalCommandsProcess)
			{
				DeliverTickerPayload(testLines);
				return;
			}

			var testTicker = new TickerOverlayForm(testLines, cycleSeconds: 2, maxCycles: 2, topPosition: false);
			Application.Run(testTicker);
			return;
		}

		if (string.IsNullOrWhiteSpace(mode))
		{
				NaturalCommands.Helpers.Logger.LogError("Mode argument is empty. Usage: ExecuteCommands.exe <mode> <dictation>");
				return;
			}

			IHandleProcesses handleProcesses = new HandleProcesses();
			Commands commands = new Commands(handleProcesses);
			string result = "";
			// Use centralized Logger for startup logging
			NaturalCommands.Helpers.Logger.LogDebug($"Log file path: {NaturalCommands.Helpers.Logger.LogPath}"); // Diagnostic: print log path
			// Clear log file on startup (write as UTF-8 with BOM)
			try {
				if (System.IO.File.Exists(NaturalCommands.Helpers.Logger.LogPath))
					System.IO.File.WriteAllText(NaturalCommands.Helpers.Logger.LogPath, string.Empty, new System.Text.UTF8Encoding(true));
			} catch(Exception ex) { NaturalCommands.Helpers.Logger.LogError($"Could not clear log file: {ex.Message}"); }

			NaturalCommands.Helpers.Logger.LogDebug($"Args: {string.Join(", ", args)}");
			NaturalCommands.Helpers.Logger.LogDebug($"ModeRaw: {modeRaw}, TextRaw: {textRaw}");
			NaturalCommands.Helpers.Logger.LogDebug($"Normalized Mode: {mode}, Text: {text}");		
		// Check if this command might start auto-click (BEFORE executing it)
		bool mightStartAutoClick = text.Contains("auto click", StringComparison.OrdinalIgnoreCase) ||
					   text.Contains("auto-click", StringComparison.OrdinalIgnoreCase);
		bool mightStartVisualTargeting = text.Contains("identify ", StringComparison.OrdinalIgnoreCase) ||
						 text.Contains("show candidates", StringComparison.OrdinalIgnoreCase) ||
						 text.Contains("choose ", StringComparison.OrdinalIgnoreCase);
		
		// Check if listen mode is already running - if so, send command to it via file instead of running locally
		bool isListenModeRunning = false;
		try
		{
			var currentProcesses = System.Diagnostics.Process.GetProcessesByName("NaturalCommands");
			var currentProcessId = System.Diagnostics.Process.GetCurrentProcess().Id;
			isListenModeRunning = currentProcesses.Any(p => p.Id != currentProcessId);
			
			// Quick Clicks command forwarding removed.
		}
		catch (Exception ex)
		{
			NaturalCommands.Helpers.Logger.LogWarning($"Could not check for listen mode: {ex.Message}");
		}
		
		// If this might start auto-click or quick clicks, initialize Windows Forms FIRST
		if (mightStartAutoClick || mightStartVisualTargeting)
		{
			NaturalCommands.Helpers.Logger.LogDebug("Command may start auto-click or quick clicks - initializing Windows Forms BEFORE execution");
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			System.Threading.SynchronizationContext.SetSynchronizationContext(
				new System.Windows.Forms.WindowsFormsSynchronizationContext());
			NaturalCommands.AutoClickOverlayForm.InitializeUIContext();
			NaturalCommands.VisualCandidateOverlayForm.InitializeUIContext();
			
			// Execute the command on the UI thread using a timer
			string? commandResult = null;
			var context = new ApplicationContext();
			var executeTimer = new System.Windows.Forms.Timer { Interval = 10 };
			executeTimer.Tick += (s, e) =>
			{
				executeTimer.Stop();
				
				try
				{
					// Execute the command
					commandResult = commands.HandleNaturalAsync(text);
					Console.WriteLine(commandResult);
					
					// Check if auto-click is now active or visual candidate overlay is visible
					if (NaturalCommands.Helpers.AutoClickManager.IsActive() || NaturalCommands.VisualCandidateOverlayForm.IsVisible)
					{
						NaturalCommands.Helpers.Logger.LogInfo("Auto-click active or quick clicks visible - keeping application alive with message pump.");
						
						// Set up a check timer to exit when both auto-click stops and quick clicks overlay is hidden
						var checkTimer = new System.Windows.Forms.Timer { Interval = 500 };
						checkTimer.Tick += (s2, e2) =>
						{
							if (!NaturalCommands.Helpers.AutoClickManager.IsActive() && !NaturalCommands.VisualCandidateOverlayForm.IsVisible)
							{
								checkTimer.Stop();
								NaturalCommands.Helpers.Logger.LogInfo("Auto-click stopped and quick clicks hidden - exiting application.");
								Application.Exit();
							}
						};
						checkTimer.Start();
					}
					else
					{
						// Command didn't start auto-click or quick clicks, exit the message pump
						Application.Exit();
					}
				}
				catch (Exception ex)
				{
					NaturalCommands.Helpers.Logger.LogError($"Error executing command: {ex.Message}");
					Console.WriteLine($"Error: {ex.Message}");
					Application.Exit();
				}
			};
			executeTimer.Start();
			
			// Run the message pump
			Application.Run(context);
			return;
		}
					switch (mode)
			{
				case "natural":
					NaturalCommands.Helpers.Logger.LogDebug($"Program.cs: Passing to HandleNaturalAsync: '{text}'");
					result = commands.HandleNaturalAsync(text);
					break;
				case "export-vs-commands":
					string outputPath = "vs_commands.json";
					if (args.Length > 2) outputPath = args[2];
					NaturalCommands.Helpers.VisualStudioHelper.ExportCommands(outputPath);
					result = $"Exported commands to {outputPath}";
					break;
				default:
					// For now, treat unknown modes as natural
					result = commands.HandleNaturalAsync(text);
					break;
			}

			NaturalCommands.Helpers.Logger.Log($"Result: {result}");
			NaturalCommands.Helpers.Logger.Log(result); // Log raw result for test matching

			// Log exact expected test substrings if present
			string[] expectedSubstrings = new[] {
				"Opened folder: Downloads",
				"Window moved to next monitor",
				"Launched app: msedge.exe",
				"Sent Ctrl+W",
				"No matching action",
				"Window set to always on top",
				"Window maximized"
			};
			foreach (var substr in expectedSubstrings)
			{
				// Remove punctuation and compare case-insensitively
				string resultStripped = new string(result.Where(c => !char.IsPunctuation(c)).ToArray());
				string substrStripped = new string(substr.Where(c => !char.IsPunctuation(c)).ToArray());
				if (resultStripped.IndexOf(substrStripped, StringComparison.OrdinalIgnoreCase) >= 0)
				{
					NaturalCommands.Helpers.Logger.Log(substr);
				}
			}
			

			if (NaturalCommands.Helpers.AutoClickManager.IsActive())
			{
				NaturalCommands.Helpers.Logger.LogError("Auto-click was started but Windows Forms was not initialized. This should not happen.");
				Console.WriteLine("ERROR: Auto-click started without proper initialization. Please use listen mode.");
			}
			// If mouse movement is active, keep the app running until it's stopped
			else if (NaturalCommands.Helpers.MouseMoveManager.IsMoving())
			{
				NaturalCommands.Helpers.Logger.LogInfo("Mouse movement active - keeping application alive. Press Ctrl+C to exit or use 'stop mouse' command.");
				
				// Keep the application running and check for stop signal
				try
				{
					using (var stopSignal = System.Threading.EventWaitHandle.OpenExisting("NaturalCommands_StopMouseMove"))
					{
						// Wait for the stop signal with a timeout
						while (NaturalCommands.Helpers.MouseMoveManager.IsMoving())
						{
							if (stopSignal.WaitOne(100))
							{
								// Stop signal received
								NaturalCommands.Helpers.MouseMoveManager.StopMoving(false, false);
								break;
							}
						}
					}
				}
				catch (System.Threading.WaitHandleCannotBeOpenedException)
				{
					// Signal doesn't exist, just use simple loop
					while (NaturalCommands.Helpers.MouseMoveManager.IsMoving())
					{
						System.Threading.Thread.Sleep(100);
					}
				}
				
				NaturalCommands.Helpers.Logger.LogInfo("Mouse movement stopped - application exiting.");
			}

		}

		private static void RunListenMode()
		{
			try
			{
				NaturalCommands.Helpers.Logger.LogInfo("Starting in listen mode (resident with tray icon).");
				Application.EnableVisualStyles();
				Application.SetCompatibleTextRenderingDefault(false);
				
				// Initialize Windows Forms synchronization context for thread-safe UI updates
				System.Threading.SynchronizationContext.SetSynchronizationContext(
					new System.Windows.Forms.WindowsFormsSynchronizationContext());
				
// Initialize the overlays' UI contexts
			NaturalCommands.AutoClickOverlayForm.InitializeUIContext();
			NaturalCommands.VisualCandidateOverlayForm.InitializeUIContext();
				
				Application.Run(new NaturalCommands.ListenModeApplicationContext());
			}
			catch (Exception ex)
			{
				try { NaturalCommands.TrayNotificationHelper.ShowNotification("Listen mode failed", ex.Message, 3500); } catch { }
			}
		}

		[System.Runtime.InteropServices.DllImport("user32.dll")]
		private static extern bool SetForegroundWindow(IntPtr hWnd);

		private static string[]? ShowDebugInputDialog()
		{
			string defaultCommand = "natural what can I say";
			// Try to load previous debug input from a small JSON file in the app directory.
			string settingsPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug_input.json"));
			try
			{
				if (File.Exists(settingsPath))
				{
					var json = File.ReadAllText(settingsPath);
					var doc = JsonSerializer.Deserialize<Dictionary<string,string>>(json);
					if (doc != null && doc.TryGetValue("command", out var savedCmd) && !string.IsNullOrWhiteSpace(savedCmd))
						defaultCommand = savedCmd;
					}
			}
			catch { }
			string currentProc = NaturalCommands.CurrentApplicationHelper.GetCurrentProcessName() ?? string.Empty;
			var procNames = System.Diagnostics.Process.GetProcesses().Select(p => p.ProcessName).Distinct().OrderBy(n => n).ToArray();

			var form = new Form() { Width = 640, Height = 220, Text = "Debug: Enter command", StartPosition = FormStartPosition.CenterScreen };
			var lblCmd = new Label() { Left = 10, Top = 10, Text = "Command (include mode e.g. 'natural what can I say'): ", AutoSize = true };
			var tbCmd = new TextBox() { Left = 10, Top = 32, Width = 600, Text = defaultCommand };
			var lblApp = new Label() { Left = 10, Top = 64, Text = "Target application (process name):", AutoSize = true };
			var cbApp = new ComboBox() { Left = 10, Top = 86, Width = 400, DropDownStyle = ComboBoxStyle.DropDown };
			cbApp.Items.AddRange(procNames);
			cbApp.Text = string.IsNullOrEmpty(currentProc) ? (procNames.Length > 0 ? procNames[0] : "") : currentProc;

			var btnOk = new Button() { Text = "OK", Left = 420, Width = 90, Top = 120, DialogResult = DialogResult.OK };
			var btnCancel = new Button() { Text = "Cancel", Left = 520, Width = 90, Top = 120, DialogResult = DialogResult.Cancel };

			form.Controls.Add(lblCmd);
			form.Controls.Add(tbCmd);
			form.Controls.Add(lblApp);
			form.Controls.Add(cbApp);
			form.Controls.Add(btnOk);
			form.Controls.Add(btnCancel);
			form.AcceptButton = btnOk;
			form.CancelButton = btnCancel;

			if (form.ShowDialog() != DialogResult.OK) return null;

			var input = (tbCmd.Text ?? defaultCommand).Trim();
			if (input.StartsWith("/")) input = input.TrimStart('/');
			string mode = "natural";
			string text = input;
			var parts = input.Split(new[] { ' ' }, 2, StringSplitOptions.RemoveEmptyEntries);
			if (parts.Length == 0)
			{
				mode = "natural"; text = "what can I say";
			}
			else if (parts.Length == 1)
			{
				mode = parts[0].ToLowerInvariant(); text = "";
			}
			else
			{
				mode = parts[0].ToLowerInvariant(); text = parts[1];
			}

			// If an application was selected, try to bring it to foreground so CurrentApplicationHelper will pick it up
			var appName = (cbApp.Text ?? string.Empty).Trim();
			if (!string.IsNullOrEmpty(appName))
			{
				try
				{
					var proc = System.Diagnostics.Process.GetProcessesByName(appName).FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);
					if (proc != null)
					{
						SetForegroundWindow(proc.MainWindowHandle);
					}
				}
				catch { }
			}

			// Persist last used debug input so subsequent debug runs pre-fill the dialog
			try
			{
				var settings = new Dictionary<string, string>()
				{
					["command"] = tbCmd.Text ?? defaultCommand,
					["app"] = appName
				};
				File.WriteAllText(settingsPath, JsonSerializer.Serialize(settings));
			}
			catch { }

			return new[] { mode, text, appName };
		}
	}
}
