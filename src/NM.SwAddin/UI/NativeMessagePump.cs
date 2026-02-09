using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Pumps Win32 messages using the raw PeekMessage/TranslateMessage/DispatchMessage
    /// loop instead of Application.DoEvents(). This avoids WinForms' internal message
    /// routing which can interfere with SolidWorks' keyboard processing.
    ///
    /// When Application.DoEvents() is used inside a SolidWorks add-in, it takes over
    /// message pumping from SolidWorks' own message loop. WinForms' message routing
    /// can prevent keyboard messages from reaching SolidWorks' native controls
    /// (property manager, feature tree, etc.). Using a raw Win32 pump preserves
    /// normal message flow to SolidWorks' windows.
    /// </summary>
    internal static class NativeMessagePump
    {
        private const int PM_REMOVE = 0x0001;
        private const int QS_ALLINPUT = 0x04FF;
        private const int WAIT_TIMEOUT = 258;

        [StructLayout(LayoutKind.Sequential)]
        private struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public int pt_x;
            public int pt_y;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PeekMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool TranslateMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessage(ref MSG lpMsg);

        [DllImport("user32.dll")]
        private static extern int MsgWaitForMultipleObjectsEx(
            int nCount, IntPtr[] pHandles, int dwMilliseconds, int dwWakeMask, int dwFlags);

        /// <summary>
        /// Waits for a condition to become true while pumping messages using
        /// the raw Win32 message pump (not Application.DoEvents).
        /// </summary>
        /// <param name="condition">Function that returns true when waiting should stop.</param>
        /// <param name="pollIntervalMs">How often to check the condition (default 50ms).</param>
        public static void WaitWithPump(Func<bool> condition, int pollIntervalMs = 50)
        {
            while (!condition())
            {
                // Wait for messages or timeout
                MsgWaitForMultipleObjectsEx(0, null, pollIntervalMs, QS_ALLINPUT, 0);

                // Pump all pending messages
                MSG msg;
                while (PeekMessage(out msg, IntPtr.Zero, 0, 0, PM_REMOVE))
                {
                    TranslateMessage(ref msg);
                    DispatchMessage(ref msg);
                }
            }
        }
    }
}
