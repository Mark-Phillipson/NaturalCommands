using System.Diagnostics;

namespace NaturalCommands
{
    public class HandleProcesses : IHandleProcesses
    {
        public void CloseAllProcesses(string processName)
        {
            var processes = Process.GetProcessesByName(processName);
            foreach (var process in processes)
            {
                process.Kill();
            }
        }
    }

}
