using DictationBoxMSP;
using SmartComponents.LocalEmbeddings;
using System.Diagnostics;
using System.Runtime.InteropServices;
using WindowsInput;
using WindowsInput.Native;
//dotnet build NaturalCommands.csproj -c Release

namespace NaturalCommands
{
	public class Commands
	{
		[DllImport("user32.dll")]
		public static extern bool ShowCursor(bool bShow);
		[DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
		public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		// Activate an application window.
		[DllImport("USER32.DLL")]
		public static extern bool SetForegroundWindow(IntPtr hWnd);
		[DllImport("user32.dll")]
		public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out uint ProcessId);

		[DllImport("user32.dll")]
		public static extern IntPtr GetForegroundWindow();



		[DllImport("user32.dll")]
		static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

		public Process? currentProcess { get; set; }
		readonly IHandleProcesses _handleProcesses;
		readonly InputSimulator inputSimulator = new InputSimulator();
		string[] arguments= Array.Empty<string>();
		readonly NaturalLanguageInterpreter _naturalInterpreter = new NaturalLanguageInterpreter();
		public Commands(IHandleProcesses handleProcesses)
		{
			_handleProcesses = handleProcesses;
		}

		/// <summary>
		/// Handles natural language commands by delegating to NaturalLanguageInterpreter
		/// </summary>
		public string HandleNaturalAsync(string text)
		{
			return _naturalInterpreter.HandleNaturalAsync(text);
		}
		
		private void UpdateTheCurrentProcess()
		{
			IntPtr hwnd = GetForegroundWindow();
			uint pid;
			GetWindowThreadProcessId(hwnd, out pid);
			currentProcess = Process.GetProcessById((int)pid);

		}

	}
}