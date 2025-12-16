using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace NaturalCommands
{
    [Flags]
    public enum HotkeyModifiers : uint
    {
        None = 0,
        Alt = 0x0001,
        Control = 0x0002,
        Shift = 0x0004,
        Win = 0x0008,
        NoRepeat = 0x4000,
    }

    public sealed class HotkeyRegistrar : IDisposable
    {
        private const int WM_HOTKEY = 0x0312;

        private readonly HotkeyWindow _window;
        private readonly HashSet<int> _registeredIds = new();
        private int _nextId = 1;
        private bool _disposed;

        public event EventHandler? Activated;

        public HotkeyRegistrar()
        {
            _window = new HotkeyWindow();
            _window.HotkeyPressed += (_, __) => Activated?.Invoke(this, EventArgs.Empty);
        }

        public bool TryRegister(HotkeyModifiers modifiers, Keys key)
        {
            if (_disposed) return false;
            int id = _nextId++;
            bool ok = RegisterHotKey(_window.Handle, id, (uint)modifiers, (uint)key);
            if (!ok) return false;
            _registeredIds.Add(id);
            return true;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                foreach (var id in _registeredIds)
                {
                    try { UnregisterHotKey(_window.Handle, id); } catch { }
                }
                _registeredIds.Clear();
            }
            catch { }

            try { _window.Dispose(); } catch { }
        }

        private sealed class HotkeyWindow : NativeWindow, IDisposable
        {
            public event EventHandler? HotkeyPressed;

            public HotkeyWindow()
            {
                CreateParams cp = new CreateParams
                {
                    Caption = "NaturalCommandsHotkeyWindow",
                    X = 0,
                    Y = 0,
                    Height = 0,
                    Width = 0,
                    Style = 0,
                };
                CreateHandle(cp);
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_HOTKEY)
                {
                    try { HotkeyPressed?.Invoke(this, EventArgs.Empty); } catch { }
                }
                base.WndProc(ref m);
            }

            public void Dispose()
            {
                try { DestroyHandle(); } catch { }
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    }
}
