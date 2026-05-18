using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;

namespace oPenEfficiency.Utils
{
    public static class CommonUtils
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out POINT lpPoint);

        public static Point GetMousePosition()
        {
            if (GetCursorPos(out POINT lpPoint))
            {
                return new Point(lpPoint.X, lpPoint.Y);
            }
            return new Point(0, 0);
        }

        public static List<string> SafeGetFiles(string rootPath, string searchPattern = "*.*")
        {
            var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var pending = new Stack<string>();
            pending.Push(rootPath);

            while (pending.Count > 0)
            {
                string path = pending.Pop();
                try
                {
                    foreach (var file in Directory.GetFiles(path, searchPattern))
                    {
                        files.Add(Path.GetFullPath(file));
                    }
                    
                    foreach (var subdir in Directory.GetDirectories(path))
                    {
                        pending.Push(subdir);
                    }
                }
                catch (UnauthorizedAccessException) { }
                catch (PathTooLongException) { }
                catch (Exception) { }
            }

            return new List<string>(files);
        }
    }
}
