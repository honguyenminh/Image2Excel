using Image2Excel.Core;

namespace Image2Excel.Legacy;

public static class Program
{
    private static void Main(string[] args)
    {
        var version = Core.Metadata.Version.Default;
        Console.WriteLine("Image2Excel.Legacy " + version.VersionString);
        Console.WriteLine("GitHub: https://github.com/honguyenminh/Image2Excel.Legacy");
        if (version.IsPreRelease)
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

        bool quantizeImage = false; // TODO: add option to switch this
        try
        {
            PackageFileCreator xl = new(consoleHandler.ImagePath);
            
            Console.Write("[Info]  Writing metadata...");
            xl.WriteMetadata(version);
            Console.WriteLine(" Done.");

            Console.Write("[Info*] Temp folder is at: ");
            Console.WriteLine(xl.TempDirectoryPath);

            WriteStylesAgain: // I am sorry
            if (quantizeImage)
            {
                Console.Write("[Info]  Quantizing color...");
                // TODO: add image quantizing here
                Console.WriteLine(" Done.");
            }

            Console.Write("[Info]  Writing styles...");
            try
            {
                xl.WriteStyles();
                Console.WriteLine(" Done.");
            }
            catch (ArgumentOutOfRangeException e)
            {
                // Too many colors
                Console.WriteLine();
                Console.WriteLine("[Error] Image file exceeded Excel's limitations.");
                Console.Write("        ");
                Console.WriteLine(e.Message);

                Console.Write("Try again with color quantizing? [Y/n]: ");
                string ans = Console.ReadLine()!.ToLower().Trim();
                if (ans != "y" && ans != "")
                {
                    Console.WriteLine("[DONE]  Operation aborted. Have a nice day owo");
                    return;
                }

                quantizeImage = true;
                goto WriteStylesAgain;
            }

            Console.Write("[Info]  Writing sheet...");
            xl.WriteSheet();
            Console.WriteLine(" Done.");

            Console.Write("[Info]  Compressing to package...");
            xl.Save(consoleHandler.OutputPath);
            Console.WriteLine(" Done.");
            
            Console.Write("[Info]  Cleaning up temp folder...");
            xl.Dispose();
            Console.WriteLine(" Done.");

            Console.WriteLine("[DONE]  All done, my master! owo");
        }
        catch (Exception e)
        {
            Console.WriteLine();
            Console.WriteLine("[Error] Oopsie! Unknown exception thrown.");
            Console.WriteLine("[Error] We don't know what went wrong yet, but here's the details:");
            Console.Write("        ");
            Console.WriteLine(e.Message);
        }
    }
}