using System;
using System.IO;

namespace Image2Excel
{
    public class Program
    {
        public const bool IS_PREPRODUCTION = true;
        public const int MAJOR_VERSION = 0;
        public const int MINOR_VERSION = 1;
        public static readonly string VERSION = $"v{MAJOR_VERSION}.{MINOR_VERSION}" + (IS_PREPRODUCTION ? "pre" : "");

        static void Main(string[] args)
        {
            Console.WriteLine("Image2Excel " + VERSION);
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
                switch (Console.ReadLine().ToLower())
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

            Console.WriteLine(parser.ImagePath);
            Console.WriteLine(parser.OutputPath);
        }

        /// <summary>
        /// Parse the arguments list and handle result
        /// </summary>
        /// <param name="args">Arguments list</param>
        /// <param name="parser"><see cref="ArgParser"/> parser object</param>
        /// <returns><see langword="true"/> if parsed successfully, <see langword="false"/> otherwise</returns>
        static bool TryParseArgumentHandler(string[] args, ref ArgParser parser)
        {
            ParseResult result = parser.TryParse(args);
            switch (result)
            {
                case ParseResult.Success:
                    return true;
                case ParseResult.Failed:
                    Console.WriteLine("Invalid argument(s). Check your command again maybe?");
                    Console.WriteLine($"Found {args.Length} argument(s)");
                    Console.WriteLine("Run 'Image2Excel --help' for help");
                    return false;
                case ParseResult.Help:
                    Console.WriteLine("Syntax: Image2Excel <image-path> <output-path>(optional)");
                    Console.WriteLine(" - Image2Excel is the path to the executable");
                    Console.WriteLine(" - <image-path> is the path to source image");
                    Console.WriteLine(" - <output-path> (optional) is the path to output file");
                    Console.WriteLine("   *If left out, this will be the image path with '.xlsx' appended*" + Environment.NewLine);
                    Console.WriteLine("THIS APP HAS SUPER OWO POWER");
                    return false;
                default:
                    throw new ArgumentOutOfRangeException(nameof(result), "Unknown ParseResult enum value");
            }
        }
    }
}
