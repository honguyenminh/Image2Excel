using System.Text;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Image2Excel;

public static class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Image2Excel " + VersionTags.Version);
        Console.WriteLine("GitHub: https://github.com/honguyenminh/Image2Excel");
        if (VersionTags.IsPreRelease)
        {
            Console.WriteLine("THIS IS A PRE-RELEASE, FEATURES MIGHT BE UNSTABLE!");
            Console.WriteLine("you have been warned uwu");
        }
        Console.WriteLine("---------------------------------------------------");

        // Parse arguments
        ConsoleHandler consoleHandler = new();
        bool success = consoleHandler.TryParse(args);
        if (!success) return;

        // Check if input paths are legit
        if (!File.Exists(consoleHandler.ImagePath))
        {
            Console.WriteLine("[Error] Cannot find/read image file at provided path.");
            return;
        }
        if (File.Exists(consoleHandler.OutputPath))
        {
            Console.WriteLine("[Error] Output file already existed.");
            Console.WriteLine("        Output path: " + consoleHandler.OutputPath);
            Console.Write("******  Do you want to overwrite this file? [y/N]: ");
            switch (Console.ReadLine()?.ToLower().Trim())
            {
                // Default is no
                case "":
                case "n":
                    Console.WriteLine("[Info]  Aborted operation. Have a nice day~");
                    return;
                case "y":
                    Console.WriteLine("[Info]  Don't blame me then owo. Overwriting file...");
                    break;
                default:
                    Console.WriteLine("[Error] What was that? owo");
                    Console.WriteLine("[Info]  Aborted operation. Have a nice day~");
                    return;
            }
            File.Delete(consoleHandler.OutputPath);
        }

        using Image<Rgba32> image = Image.Load<Rgba32>(consoleHandler.ImagePath);

        // TODO: add verbose toggle
        try
        {
            Console.Write("[Info]  Writing metadata...");
            using ExcelHandler xl = new();
            Console.WriteLine(" Done.");
            
            Console.Write("[Info*] Temp folder is at: ");
            Console.WriteLine(xl.TempDirectoryPath);
            
            Console.Write("[Info]  Writing styles...");
            xl.WriteStyles(image);
            Console.WriteLine(" Done.");
            
            Console.Write("[Info]  Writing sheet...");
            xl.WriteSheet(image);
            Console.WriteLine(" Done.");
            Console.Write("[Info]  Compressing to package...");
            xl.Save(consoleHandler.OutputPath);
            Console.WriteLine(" Done.");
            
            Console.WriteLine("[DONE]  All done, my master! owo");
        }
        catch (ArgumentOutOfRangeException e)
        {
            Console.WriteLine();
            Console.WriteLine("[Error] Image file exceeded Excel's limitations.");
            Console.Write("        ");
            Console.WriteLine(e.Message);
        }
        // TODO: add other errors exception handler here
    }

    public static string ToArgbHex(this Rgba32 color)
    {
        uint hexOrder = (uint)(color.B << 0 | color.G << 8 | color.R << 16 | color.A << 24);
        return hexOrder.ToString("X8");
    }
}