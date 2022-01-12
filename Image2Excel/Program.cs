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
        Console.WriteLine("GitHub: https://github.com/honguyenminh/Image2Excel" + Environment.NewLine);

        // Parse arguments
        ArgParser parser = new();
        bool success = TryParseArgumentHandler(args, ref parser);
        if (!success) return;

        // Check if input paths are legit
        if (!File.Exists(parser.ImagePath))
        {
            Console.WriteLine("Cannot find/read image file at provided path.");
            return;
        }
        if (File.Exists(parser.OutputPath))
        {
            Console.WriteLine("Output file already existed.");
            Console.WriteLine("Output path: " + parser.OutputPath);
            Console.Write("Do you want to overwrite this file? [y/N]: ");
            switch (Console.ReadLine()?.ToLower())
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
            File.Delete(parser.OutputPath);
        }

        // Process image
        Dictionary<string, int> fillStyleIdSet = new();
        StringBuilder colorStyles = new();
        using Image<Rgb24> img = Image.Load<Rgb24>("test.png");
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

    /// <summary>
    /// Parse the arguments list and handle result
    /// </summary>
    /// <param name="args">Arguments list</param>
    /// <param name="parser"><see cref="ArgParser"/> parser object</param>
    /// <returns><see langword="true"/> if parsed successfully, <see langword="false"/> otherwise</returns>
    private static bool TryParseArgumentHandler(string[] args, ref ArgParser parser)
    {
        ParseResult result = parser.TryParse(args);
        switch (result)
        {
            case ParseResult.Success:
                return true;
            case ParseResult.Failed:
                Console.WriteLine("Invalid argument(s). Check your command again maybe?");
                Console.WriteLine($"Found {args.Length} argument(s)");
                Console.WriteLine($"Run 'Image2Excel --help' for help");
                return false;
            case ParseResult.Help:
                Console.WriteLine("Syntax: <Image2Excel> <image-path> <output-path>(optional)");
                Console.WriteLine(" - <Image2Excel> is the path to the executable");
                Console.WriteLine(" - <image-path> is the path to source image");
                Console.WriteLine(" - <output-path> (optional) is the path to output file");
                Console.WriteLine("   *If left out, output path will be the image path with '.xlsx' appended*" + Environment.NewLine);
                Console.WriteLine("THIS APP HAS SUPER OWO POWER");
                return false;
            default:
                throw new ArgumentOutOfRangeException(nameof(result), "Unknown ParseResult enum value");
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