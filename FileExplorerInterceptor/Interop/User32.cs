using FileExplorerInterceptor.Interop.Delegates;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace FileExplorerInterceptor.Interop
{
    public static class User32
    {
        public static uint EVENT_OBJECT_CREATE = 0x8000;
        public static uint WINEVENT_OUTOFCONTEXT = 0;
        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out IntPtr processId);

        [DllImport("user32.dll")]
        public static extern uint RealGetWindowClass(IntPtr hwnd, StringBuilder pszType, uint cchType);


        public static IntPtr HWND_BOTTOM = new IntPtr(1);
        public static IntPtr HWND_TOP = new IntPtr(0);
        public static uint SWP_NOSIZE = 0x0001;
        public static uint SWP_NOMOVE = 0x0002;
        public static uint SWP_NOACTIVATE = 0x0010;
        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int Y, int cx, int cy, uint wFlags);
    }
}
