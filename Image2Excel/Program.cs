using System.Text;
using System.Reflection;
using System.Resources;
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

        // Process image
        Dictionary<string, int> fillStyleIdSet = new();
        StringBuilder colorStyles = new();
        using Image<Rgb24> img = Image.Load<Rgb24>("test.png"); // TODO: replace this
        for (int y = 0; y < img.Height; y++)
        {
            if (fillStyleIdSet.Count > 255)
            {
                Console.WriteLine("Too many colors. Max is 256 (blame Excel owo)");
                return;
            }
    
            Span<Rgb24> row = img.GetPixelRowSpan(y);
            for (int x = 0; x < img.Width; x++)
            {
                if (fillStyleIdSet.ContainsKey(ToHex(row[x]))) continue;
                fillStyleIdSet.Add(ToHex(row[x]), fillStyleIdSet.Count);
                colorStyles.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"");
                colorStyles.Append(ToHex(row[x]));
                colorStyles.Append("\"/><bgColor indexed=\"64\"/></patternFill></fill>");
            }
        }
    }

    private static string ToHex(Rgb24 color)
        => color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");

    private static string GetEmbeddedResourceContent(string resourceName)
    {
        var asm = Assembly.GetExecutingAssembly();
        var stream = asm.GetManifestResourceStream(resourceName);
        if (stream is null) 
            throw new MissingManifestResourceException("Cannot find resource " + resourceName);

        StreamReader source = new(stream);
        string fileContent = source.ReadToEnd();
        source.Dispose();
        stream.Dispose();
        return fileContent;
    }
}