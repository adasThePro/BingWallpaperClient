// Simple C# program to set the desktop wallpaper on Windows
// using SystemParametersInfoW API.
// Windows 7 or later

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

class SetWallpaper
{
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern int SystemParametersInfoW(int uAction, int uParam, string lpvParam, int fuWinIni);
    
    static bool IsWindows7()
    {
        Version version = Environment.OSVersion.Version;
        return version.Major == 6 && version.Minor == 1;
    }
    
    static int Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Error: No image path provided");
            Console.Error.WriteLine("Usage: SetWallpaper.exe <path_to_image>");
            return 1;
        }
        
        string imagePath = args[0];
        
        if (!File.Exists(imagePath))
        {
            Console.Error.WriteLine("Error: File not found - " + imagePath);
            return 2;
        }
        
        string ext = Path.GetExtension(imagePath).ToLower();
        if (ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".bmp")
        {
            Console.Error.WriteLine("Error: Unsupported image format. Use JPG, PNG, or BMP");
            return 3;
        }
        
        imagePath = Path.GetFullPath(imagePath);
        
        if (IsWindows7() && ext != ".bmp")
        {
            try
            {
                string tempPath = Path.Combine(Path.GetTempPath(), "wallpaper.bmp");
                Image img = Image.FromFile(imagePath);
                img.Save(tempPath, ImageFormat.Bmp);
                img.Dispose();
                imagePath = tempPath;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: Failed to convert image - " + ex.Message);
                return 5;
            }
        }
        
        int result = SystemParametersInfoW(20, 0, imagePath, 3);
        
        if (result != 0)
        {
            Console.WriteLine("Wallpaper set successfully: " + imagePath);
            return 0;
        }
        else
        {
            Console.Error.WriteLine("Error: Failed to set wallpaper");
            return 4;
        }
    }
}