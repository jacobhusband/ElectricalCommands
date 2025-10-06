// WindowOrchestrator.cs
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCADCleanupTool
{
    internal static class WindowOrchestrator
    {
        // --- Win32 interop ---
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(
            IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        private static readonly IntPtr HWND_TOP = IntPtr.Zero;
        private const int SW_RESTORE = 9;
        private const int SW_MINIMIZE = 6;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private static Rectangle GetRect(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return Rectangle.Empty;
            if (!GetWindowRect(hwnd, out var r)) return Rectangle.Empty;
            return Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom);
        }

        private static Rectangle Intersect(Rectangle a, Rectangle b)
        {
            int x1 = Math.Max(a.Left, b.Left);
            int y1 = Math.Max(a.Top, b.Top);
            int x2 = Math.Min(a.Right, b.Right);
            int y2 = Math.Min(a.Bottom, b.Bottom);
            if (x2 <= x1 || y2 <= y1) return Rectangle.Empty;
            return Rectangle.FromLTRB(x1, y1, x2, y2);
        }

        /// <summary>
        /// Try to find PPT's main window handle.
        /// If you already have an Interop.Application, pass its HWND via pptAppHwndHint.
        /// </summary>
        public static bool TryGetPowerPointHwnd(out IntPtr pptHwnd, IntPtr pptAppHwndHint = default)
        {
            pptHwnd = IntPtr.Zero;

            // 1) Preferred: Interop Application.HWND (pass it in)
            if (pptAppHwndHint != IntPtr.Zero)
            {
                pptHwnd = pptAppHwndHint;
                return true;
            }

            // 2) Fallback: find running POWERPNT process
            try
            {
                foreach (var p in Process.GetProcessesByName("POWERPNT"))
                {
                    if (p.MainWindowHandle != IntPtr.Zero)
                    {
                        pptHwnd = p.MainWindowHandle;
                        return true;
                    }
                }
            }
            catch { /* ignore */ }

            return false;
        }

        /// <summary>
        /// Ensure PPT is on a different monitor than AutoCAD; if not possible, minimize it for safety.
        /// </summary>
        public static void EnsureSeparationOrSafeOverlap(Editor ed, IntPtr pptHwnd, bool preferDifferentMonitor = true)
        {
            try
            {
                IntPtr acadHwnd = Application.MainWindow.Handle;
                if (acadHwnd == IntPtr.Zero || pptHwnd == IntPtr.Zero) return;

                var acadRect = GetRect(acadHwnd);
                var pptRect = GetRect(pptHwnd);

                // If we have >=2 monitors, try to move PPT to a screen that isn't hosting AutoCAD
                var screens = Screen.AllScreens;
                if (preferDifferentMonitor && screens.Length >= 2)
                {
                    var acadScreen = Screen.FromRectangle(acadRect);
                    Screen target = null;
                    foreach (var s in screens)
                    {
                        if (s.DeviceName != acadScreen.DeviceName)
                        {
                            if (target == null) target = s; // first non-ACAD screen
                            else
                            {
                                // prefer the largest working area among the non-ACAD screens
                                if (s.WorkingArea.Width * s.WorkingArea.Height >
                                    target.WorkingArea.Width * target.WorkingArea.Height)
                                    target = s;
                            }
                        }
                    }

                    if (target != null)
                    {
                        // Restore (if minimized), then move/size PPT to the target working area
                        if (IsIconic(pptHwnd)) ShowWindow(pptHwnd, SW_RESTORE);
                        var wa = target.WorkingArea;
                        SetWindowPos(pptHwnd, HWND_TOP, wa.Left, wa.Top, wa.Width, wa.Height, SWP_NOZORDER | SWP_NOACTIVATE);
                        // Give DWM a beat
                        Thread.Sleep(120);

                        // If ACAD and PPT still overlap *and* share a monitor (rare), minimize PPT
                        var newPptRect = GetRect(pptHwnd);
                        if (Screen.FromRectangle(newPptRect).DeviceName == acadScreen.DeviceName &&
                            !Intersect(newPptRect, acadRect).IsEmpty)
                        {
                            ShowWindow(pptHwnd, SW_MINIMIZE);
                            SetForegroundWindow(acadHwnd);
                        }

                        ed?.WriteMessage("\nPowerPoint arranged to a different monitor.");
                        return;
                    }
                }

                // Single-monitor OR couldn't find a different one: minimize PPT (safe overlap mode)
                ShowWindow(pptHwnd, SW_MINIMIZE);
                SetForegroundWindow(acadHwnd);
                ed?.WriteMessage("\nSingle-monitor (or no alternate screen). Using safe mode: PowerPoint minimized by default.");
            }
            catch (Exception ex)
            {
                ed?.WriteMessage($"\n[WindowOrchestrator] Arrangement warning: {ex.Message}");
            }
        }

        /// <summary>
        /// Call right BEFORE you need to drive PowerPoint (single-monitor safety).
        /// </summary>
        public static void BeginPptInteraction(IntPtr pptHwnd)
        {
            try
            {
                if (pptHwnd == IntPtr.Zero) return;
                ShowWindow(pptHwnd, SW_RESTORE);
                SetForegroundWindow(pptHwnd);
                Thread.Sleep(80);
            }
            catch { /* ignore */ }
        }

        /// <summary>
        /// Call right AFTER you finish driving PowerPoint (single-monitor safety).
        /// Brings focus back to AutoCAD and minimizes PPT.
        /// </summary>
        public static void EndPptInteraction()
        {
            try
            {
                var acadHwnd = Application.MainWindow.Handle;
                // Minimize any active PPT (if any)
                if (TryGetPowerPointHwnd(out var pptHwnd))
                    ShowWindow(pptHwnd, SW_MINIMIZE);

                if (acadHwnd != IntPtr.Zero)
                    SetForegroundWindow(acadHwnd);
            }
            catch { /* ignore */ }
        }
    }
}