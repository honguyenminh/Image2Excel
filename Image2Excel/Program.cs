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
        Console.WriteLine("---------------------------------------------------");

        // Parse arguments
        ConsoleHandler console = new();
        bool success = console.TryParse(args);
        if (!success) return;

        // Check if input paths are legit
        if (!File.Exists(console.ImagePath))
        {
            Console.WriteLine("Cannot find/read image file at provided path.");
            return;
        }
        if (File.Exists(console.OutputPath))
        {
            Console.WriteLine("Output file already existed.");
            Console.WriteLine("Output path: " + console.OutputPath);
            Console.Write("Do you want to overwrite this file? [y/N]: ");
            switch (Console.ReadLine()?.ToLower().Trim())
            {
                // Default is no
                case "":
                case "n":
                    Console.WriteLine("Aborted operation. Have a nice day~");
                    return;
                case "y":
                    Console.WriteLine("Don't blame me then owo. Overwriting file...");
                    break;
                default:
                    Console.WriteLine("What was that? owo");
                    Console.WriteLine("Aborted operation. Have a nice day~");
                    return;
            }
            File.Delete(console.OutputPath);
        }

        using Image<Rgb24> image = Image.Load<Rgb24>(console.ImagePath);

        using ExcelHandler xl = new();
        xl.WriteStyles(image);
    }

    public static string ToHex(this Rgb24 color)
    {
        return color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
    }
}