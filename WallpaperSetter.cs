using System;
using System.Runtime.InteropServices;

namespace Wallpaper
{
    [ComImport]
    [Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDesktopWallpaper
    {
        void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.LPWStr)] string wallpaper);

        [return: MarshalAs(UnmanagedType.U4)]
        uint GetMonitorDevicePathCount();

        [return: MarshalAs(UnmanagedType.LPWStr)]
        string GetMonitorDevicePathAt(uint monitorIndex);
    }

    public class Setter
    {

        public static void SetWallpaperForMonitor(int monitorIndex, string wallpaperPath)
        {
            var desktopWallpaperType = Type.GetTypeFromCLSID(new Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"));
            var desktopWallpaper = (IDesktopWallpaper)Activator.CreateInstance(desktopWallpaperType);

            uint count = desktopWallpaper.GetMonitorDevicePathCount();

            if (monitorIndex >= 0 && monitorIndex < count)
            {
                string monitorID = desktopWallpaper.GetMonitorDevicePathAt((uint)monitorIndex);
                desktopWallpaper.SetWallpaper(monitorID, wallpaperPath);
            }
            else
            {
                throw new ArgumentOutOfRangeException("Monitor index out of range");
            }
        }
    }
}