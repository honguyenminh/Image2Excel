namespace Image2Excel;

internal class ConsoleHandler
{
    public string ImagePath { get; private set; } = "";
    public string OutputPath { get; private set; } = "";

    /// <summary>
    /// Try to parse arguments from console args
    /// </summary>
    /// <param name="args">Console arguments list</param>
    /// <returns>true if parsed successfully, false otherwise</returns>
    public bool TryParse(string[] args)
    {
        #if DEBUG
        ImagePath = "demo.png";
        OutputPath = "out.xlsx";
        return true;
        #endif
        // Help
        if (args.Contains("-h") || args.Contains("--help"))
        {
            WriteHelp();
            return false;
        }
        switch (args.Length)
        {
            case 1:
                // Image path only, auto make output filename
                ImagePath = args[0];
                OutputPath = args[0] + ".xlsx";
                break;
            case 2:
                // Image and output path given
                ImagePath = args[0];
                OutputPath = args[1];
                break;
            default:
                WriteInvalidArgs(args.Length);
                return false;
        }
        return true;
    }

    // Move these into "Resource" folder if needed
    private static void WriteHelp()
    {
        Console.WriteLine("Syntax: <Image2Excel> <image-path> <output-path>(optional)");
        Console.WriteLine(" - <Image2Excel> is the path to the executable");
        Console.WriteLine(" - <image-path> is the path to source image");
        Console.WriteLine(" - <output-path> (optional) is the path to output file");
        Console.WriteLine("   *If left out, output path will be the image path with '.xlsx' appended*");
        Console.WriteLine();
        Console.WriteLine("THIS APP HAS SUPER OWO POWER");
    }
    private static void WriteInvalidArgs(int argsCount)
    {
        Console.WriteLine("Invalid argument(s). Check your command again maybe?");
        Console.WriteLine($"Found {argsCount} argument(s)");
        Console.WriteLine("Run 'Image2Excel --help' for help");
    }
}