namespace NaturalCommands.Helpers
{
    // Handles sending key sequences and typing text
    public class KeySender
    {
        // TODO: Move key sending methods here from NaturalLanguageInterpreter

        public static string SendKeys(NaturalCommands.SendKeysAction keys)
        {
            string keysText = keys.KeysText?.Trim().ToLowerInvariant() ?? string.Empty;
            NaturalCommands.Helpers.Logger.LogDebug($"KeySender.SendKeys called with: '{keysText}'");

            if (keysText == "ctrl alt tab" || keysText == "control alt tab")
            {
                // Use keybd_event to send Ctrl+Alt+Tab
                NaturalCommands.Helpers.Logger.LogDebug("Sending Ctrl+Alt+Tab key sequence.");
                // Ctrl down
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyDown(0x11); // VK_CONTROL
                // Alt down
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyDown(0x12); // VK_MENU (Alt)
                // Tab down
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyDown(0x09); // VK_TAB
                // Tab up
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyUp(0x09);
                // Alt up
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyUp(0x12);
                // Ctrl up
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyUp(0x11);
                NaturalCommands.Helpers.Logger.LogDebug("Sent Ctrl+Alt+Tab.");
                return "[KeySender.SendKeys] Sent Ctrl+Alt+Tab.";
            }
            else if (keysText == "control ," || keysText == "ctrl ,")
            {
                // Send Ctrl+Comma
                NaturalCommands.Helpers.Logger.LogDebug("Sending Ctrl+Comma key sequence.");
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyDown(0x11); // VK_CONTROL
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyDown(0xBC); // VK_OEM_COMMA
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyUp(0xBC);
                NaturalCommands.Helpers.WindowFocusHelper.SendKeyUp(0x11);
                NaturalCommands.Helpers.Logger.LogDebug("Sent Ctrl+Comma.");
                return "[KeySender.SendKeys] Sent Ctrl+Comma.";
            }
            else
            {
                NaturalCommands.Helpers.Logger.LogError($"Unsupported key sequence: '{keysText}'");
                return $"[KeySender.SendKeys] Unsupported key sequence: '{keysText}'";
            }
        }

        // Sends a shortcut string like "ctrl+alt+t" or "win+h" or "ctrl+shift+f5".
        // Supported modifiers: ctrl/control, alt, shift, win/lwin
        // Supported keys: single letters, digits, function keys (f1..f24), and common punctuation via VK codes
        public static string SendShortcut(string shortcut)
        {
            if (string.IsNullOrWhiteSpace(shortcut)) return "";
            var parts = shortcut.Split(new[] { '+', '-' }, StringSplitOptions.RemoveEmptyEntries).Select(p => p.Trim().ToLowerInvariant()).ToArray();
            var modifiers = new List<byte>();
            byte? mainKey = null;

            foreach (var p in parts)
            {
                if (p == "ctrl" || p == "control") modifiers.Add(0x11); // VK_CONTROL
                else if (p == "alt" || p == "menu") modifiers.Add(0x12); // VK_MENU
                else if (p == "shift") modifiers.Add(0x10); // VK_SHIFT
                else if (p == "win" || p == "lwin" || p == "rwin") modifiers.Add(0x5B); // VK_LWIN (use left win)
                else if (p.Length == 1 && char.IsLetterOrDigit(p[0])) mainKey = (byte)char.ToUpperInvariant(p[0]);
                else if (p.StartsWith("f") && int.TryParse(p.Substring(1), out var fn) && fn >= 1 && fn <= 24)
                {
                    mainKey = (byte)(0x70 + fn - 1); // VK_F1 = 0x70
                }
                else
                {
                    // simple punctuation, NumPad keys, arrow keys and other special cases
                    switch (p)
                    {
                        case "enter": mainKey = 0x0D; break;
                        case "esc": case "escape": mainKey = 0x1B; break;
                        case "tab": mainKey = 0x09; break;
                        case "space": mainKey = 0x20; break;
                        case ",": case "comma": mainKey = 0xBC; break; // VK_OEM_COMMA
                        case ".": case "period": case "dot": mainKey = 0xBE; break; // VK_OEM_PERIOD
                        case "backspace": mainKey = 0x08; break;
                        case "delete": case "del": mainKey = 0x2E; break; // VK_DELETE
                        case "insert": case "ins": mainKey = 0x2D; break; // VK_INSERT
                        case "home": mainKey = 0x24; break; // VK_HOME
                        case "end": mainKey = 0x23; break; // VK_END
                        case "pageup": case "pgup": mainKey = 0x21; break; // VK_PRIOR
                        case "pagedown": case "pgdn": mainKey = 0x22; break; // VK_NEXT
                        
                        // Arrow keys
                        case "up": case "uparrow": mainKey = 0x26; break; // VK_UP
                        case "down": case "downarrow": mainKey = 0x28; break; // VK_DOWN
                        case "left": case "leftarrow": mainKey = 0x25; break; // VK_LEFT
                        case "right": case "rightarrow": mainKey = 0x27; break; // VK_RIGHT
                        
                        // NumPad add/subtract (common Talon toggle keys)
                        case "add": case "numpadadd": mainKey = 0x6B; break; // VK_ADD
                        case "subtract": case "numpadsubtract": mainKey = 0x6D; break; // VK_SUBTRACT
                        
                        // Plus and Minus (OEM keys, not numpad)
                        case "plus": case "=": case "equals": mainKey = 0xBB; break; // VK_OEM_PLUS (=+ key)
                        case "minus": case "-": mainKey = 0xBD; break; // VK_OEM_MINUS (-_ key)
                        
                        // Additional punctuation
                        case "[": case "openbracket": mainKey = 0xDB; break; // VK_OEM_4
                        case "]": case "closebracket": mainKey = 0xDD; break; // VK_OEM_6
                        case ";": case "semicolon": mainKey = 0xBA; break; // VK_OEM_1
                        case "'": case "quote": mainKey = 0xDE; break; // VK_OEM_7
                        case "/": case "slash": mainKey = 0xBF; break; // VK_OEM_2
                        case "\\": case "backslash": mainKey = 0xDC; break; // VK_OEM_5
                        case "`": case "backtick": case "grave": mainKey = 0xC0; break; // VK_OEM_3
                        
                        default:
                            // unknown token
                            NaturalCommands.Helpers.Logger.LogDebug($"SendShortcut: unknown token '{p}' in '{shortcut}'");
                            break;
                    }
                }
            }

            // Press modifiers
            foreach (var m in modifiers) WindowFocusHelper.SendKeyDown(m);
            if (mainKey.HasValue) WindowFocusHelper.SendKeyDown(mainKey.Value);
            if (mainKey.HasValue) WindowFocusHelper.SendKeyUp(mainKey.Value);
            // Release modifiers in reverse order
            for (int i = modifiers.Count - 1; i >= 0; i--) WindowFocusHelper.SendKeyUp(modifiers[i]);

            return $"[KeySender.SendShortcut] Sent '{shortcut}'";
        }
    }
}
