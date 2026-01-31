using System;
using System.Windows.Forms;
using DictationBoxMSP;

namespace NaturalCommands.Helpers
{
    public static class VoiceDictationHelper
    {
        /// <summary>
        /// Callback delegate for processing commands sent from the voice dictation form.
        /// </summary>
        public delegate void CommandHandler(string commandText);

        /// <summary>
        /// Shows the voice dictation form. Commands are processed via the onCommandSent callback
        /// while the form stays open for additional commands. The form closes when the user clicks Cancel.
        /// </summary>
        /// <param name="timeoutMs">Auto-submit timeout (0 = disabled)</param>
        /// <param name="onCommandSent">Callback invoked when user sends a command</param>
        public static void ShowVoiceDictation(int timeoutMs, CommandHandler? onCommandSent)
        {
            try
            {
                using (var frm = new VoiceDictationForm(timeoutMs, autoStartDictation: true))
                {
                    if (onCommandSent != null)
                    {
                        frm.CommandSent += (sender, commandText) =>
                        {
                            try
                            {
                                onCommandSent(commandText);
                            }
                            catch { }
                        };
                    }

                    frm.ShowDialog();
                    // Form stays open until user clicks Cancel
                }
            }
            catch { }
        }

        // Legacy overload for backward compatibility - returns the last command text on cancel
        // This is kept for any code that still expects the old single-command behavior
        public static string? ShowVoiceDictation(int timeoutMs = 0)
        {
            string? result = null;
            try
            {
                using (var frm = new VoiceDictationForm(timeoutMs, autoStartDictation: true))
                {
                    var dr = frm.ShowDialog();
                    if (dr == DialogResult.OK)
                        result = frm.ResultText;
                }
            }
            catch { }
            return result;
        }
    }
}
