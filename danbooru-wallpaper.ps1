# Compile C# library
if (-not ([Type]::GetType('Wallpaper.Setter'))) {
    $code = @"
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;

    namespace Wallpaper
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        /// <summary>
        /// This enumeration is used to set and get slideshow options.
        /// </summary> 
        public enum DesktopSlideshowOptions
        {
            ShuffleImages = 0x01,     // When set, indicates that the order in which images in the slideshow are displayed can be randomized.
        }


        /// <summary>
        /// This enumeration is used by GetStatus to indicate the current status of the slideshow.
        /// </summary>
        public enum DesktopSlideshowState
        {
            Enabled = 0x01,
            Slideshow = 0x02,
            DisabledByRemoteSession = 0x04,
        }


        /// <summary>
        /// This enumeration is used by the AdvanceSlideshow method to indicate whether to advance the slideshow forward or backward.
        /// </summary>
        public enum DesktopSlideshowDirection
        {
            Forward = 0,
            Backward = 1,
        }

        /// <summary>
        /// This enumeration indicates the wallpaper position for all monitors. (This includes when slideshows are running.)
        /// The wallpaper position specifies how the image that is assigned to a monitor should be displayed.
        /// </summary>
        public enum DesktopWallpaperPosition
        {
            Center = 0,
            Tile = 1,
            Stretch = 2,
            Fit = 3,
            Fill = 4,
            Span = 5,
        }

        [ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDesktopWallpaper
        {
            void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.LPWStr)] string wallpaper);
            [return: MarshalAs(UnmanagedType.LPWStr)]
            string GetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID);

            /// <summary>
            /// Gets the monitor device path.
            /// </summary>
            /// <param name="monitorIndex">Index of the monitor device in the monitor device list.</param>
            /// <returns></returns>
            [return: MarshalAs(UnmanagedType.LPWStr)]
            string GetMonitorDevicePathAt(uint monitorIndex);

            /// <summary>
            /// Gets number of monitor device paths.
            /// </summary>
            /// <returns></returns>
            [return: MarshalAs(UnmanagedType.U4)]
            uint GetMonitorDevicePathCount();

            [return: MarshalAs(UnmanagedType.Struct)]
            Rect GetMonitorRECT([MarshalAs(UnmanagedType.LPWStr)] string monitorID);

            void SetBackgroundColor([MarshalAs(UnmanagedType.U4)] uint color);
            [return: MarshalAs(UnmanagedType.U4)]
            uint GetBackgroundColor();

            void SetPosition([MarshalAs(UnmanagedType.I4)] DesktopWallpaperPosition position);
            [return: MarshalAs(UnmanagedType.I4)]
            DesktopWallpaperPosition GetPosition();

            void SetSlideshow(IntPtr items);
            IntPtr GetSlideshow();

            void SetSlideshowOptions(DesktopSlideshowDirection options, uint slideshowTick);
            [PreserveSig]
            uint GetSlideshowOptions(out DesktopSlideshowDirection options, out uint slideshowTick);

            void AdvanceSlideshow([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.I4)] DesktopSlideshowDirection direction);

            DesktopSlideshowDirection GetStatus();

            bool Enable();
        }

        /// <summary>
        /// CoClass DesktopWallpaper
        /// </summary>
        [ComImport, Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD")]
        public class DesktopWallpaperClass
        {
        }

        public class Utility
        {
            static int GCD(int a, int b)
            {
                int Remainder;

                while( b != 0 )
                {
                    Remainder = a % b;
                    a = b;
                    b = Remainder;
                }

                return a;
            }


            public static string AspectRatio(int x, int y) 
            {
                return string.Format("{0}:{1}",x/GCD(x,y), y/GCD(x,y));
            }
        }

        public class Setter
        {   
            public static IDesktopWallpaper getDesktopWallpaperInterface() 
            {
                var desktopWallpaperType = Type.GetTypeFromCLSID(new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"));
                return (IDesktopWallpaper)Activator.CreateInstance(desktopWallpaperType);
            }

            public static void SetWallpaperForMonitor(int monitorIndex, string wallpaperPath)
            {
                var desktopWallpaper = getDesktopWallpaperInterface();
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

            public static int GetWidth(int monitorIndex)
            {
                var desktopWallpaper = getDesktopWallpaperInterface();
                uint count = desktopWallpaper.GetMonitorDevicePathCount();

                if (monitorIndex >= 0 && monitorIndex < count)
                {
                    string monitorID = desktopWallpaper.GetMonitorDevicePathAt((uint)monitorIndex);
                    return System.Math.Abs(desktopWallpaper.GetMonitorRECT(monitorID).Left - desktopWallpaper.GetMonitorRECT(monitorID).Right);
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Monitor index out of range");
                }
            }

            public static int GetHeight(int monitorIndex)
            {
                var desktopWallpaper = getDesktopWallpaperInterface();
                uint count = desktopWallpaper.GetMonitorDevicePathCount();

                if (monitorIndex >= 0 && monitorIndex < count)
                {
                    string monitorID = desktopWallpaper.GetMonitorDevicePathAt((uint)monitorIndex);
                    return System.Math.Abs(desktopWallpaper.GetMonitorRECT(monitorID).Top - desktopWallpaper.GetMonitorRECT(monitorID).Bottom);
                }
                else
                {
                    throw new ArgumentOutOfRangeException("Monitor index out of range");
                }
            }

            public static uint CountMonitors() {
                var desktopWallpaper = getDesktopWallpaperInterface();
                return desktopWallpaper.GetMonitorDevicePathCount();
            }
        }
    }
"@

    # Set the wallpaper for a specific monitor using C# and COM Interface
    try {
        # Try to load the DLL
        Add-Type -TypeDefinition $code
    }
    catch {
        # If there's an error, show the details
        Write-Host "Error loading the DLL: $($_.Exception.Message)"

        # Display loader exceptions
        if ($_.Exception -is [System.Reflection.ReflectionTypeLoadException]) {
            $loaderExceptions = $_.Exception.LoaderExceptions
            foreach ($ex in $loaderExceptions) {
                Write-Host "Loader exception: $ex.Message"
            }
        }
    }
}

$numberOfMonitors = [Wallpaper.Setter]::CountMonitors()
$sources = "https://konachan.com/post.json", "https://danbooru.donmai.us/posts.json"

for ($monitor = 0; $monitor -lt $numberOfMonitors; $monitor++) {
    # Configuration
    $width = [Wallpaper.Setter]::GetWidth($monitor)
    $height = [Wallpaper.Setter]::GetHeight($monitor)
    # $ratio = [Wallpaper.Utility]::AspectRatio($width, $height) = Unsuported by Konachan
    $limit = 100
    $tags = "width:>$width height:>$height"

    $baseUrl = $sources | Get-Random
    $url = $baseUrl+"?tags=$tags&limit=$limit"

    $headers = @{
        "User-Agent" = "Other"
    }


    # GET for wallpapers list
    try {
        $response = Invoke-RestMethod -Uri $url -Method GET -Headers $headers
    }
    catch {
        Write-Host "Error fetching data: $($_.Exception.Message)"
        exit
    }

    # Random post from response
    $threshold = 24MB
    $randomElement = $response | Where-Object { $_.file_size -lt $threshold } | Get-Random

    # Get the image
    $fileUrl = $randomElement.file_url
    $fileName = [System.IO.Path]::GetFileName($fileUrl)
    $fileExt = [System.IO.Path]::GetExtension($fileName)
    $imagePath = "$env:TEMP\wallpaper$monitor$fileExt"

    $response = Invoke-WebRequest -Uri $fileUrl -OutFile $imagePath -Headers $headers

    if (Test-Path $imagePath) {
        $fileInfo = Get-Item $imagePath
        if ($fileInfo.Length -gt 0) {
            Write-Host "File downloaded successfully: $imagePath"
            
            # Set the wallpaper
            [Wallpaper.Setter]::SetWallpaperForMonitor($monitor, $imagePath)
            Write-Host "Wallpaper for monitor $monitor set from: $fileUrl"
            Write-Host "Resolution $width x $height"
        } else {
            Write-Host "File exists but is empty. Download may have failed."
        }
    } else {
        Write-Host "File not downloaded or path is incorrect: $imagePath"
    }
}