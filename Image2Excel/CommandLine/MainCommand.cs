using System.Data;
using Image2Excel.Core;
using Microsoft.Extensions.Logging;

namespace Image2Excel.CommandLine;

internal class MainCommand
{
    public static void Command(MainParams p, ConsoleLogger logger, Version version)
    {
        logger.LowestLogLevel = p.Silent ? LogLevel.Warning : LogLevel.Information;
        logger.LogInformation($"Image2Excel {version.VersionString}");
#if DEBUG
        Console.WriteLine("--------------------DEBUG BUILD--------------------");
#endif
        logger.LogInformation("GitHub: https://github.com/honguyenminh/Image2Excel");
        if (version.IsPreRelease)
        {
            logger.LogWarning("THIS IS A PRE-RELEASE, FEATURES MIGHT BE UNSTABLE!");
            Console.WriteLine("you have been warned uwu");
        }
        if (!p.Silent) Console.WriteLine("---------------------------------------------------");

        p.OutputPath ??= p.InputPath + ".xlsx"; // if output path not specified
        if (!File.Exists(p.InputPath))
        {
            logger.LogError("Cannot find/read image file at provided path");
            return;
        }

        if (File.Exists(p.OutputPath))
        {
            if (!p.Force)
            {
                logger.LogError($"Output file already existed at {p.OutputPath}");
                Console.Write("******  Do you want to overwrite output file? [y/N]: ");
                switch (Console.ReadLine()?.ToLower().Trim())
                {
                    case "y":
                        logger.LogWarning("Don't blame me then owo. Overwriting file...");
                        break;
                    // Default is no
                    default:
                        logger.LogInformation("Aborted operation. Have a nice day~");
                        return;
                }
            }
            try
            {
                File.Delete(p.OutputPath);
            }
            catch (UnauthorizedAccessException e)
            {
                logger.LogError("Can't overwrite output file, not enough permission or is read-only");
                logger.Log(LogLevel.None, $"Error message: {e.Message}");
            }
        }

        PackageFileCreator package;
        try
        {
            package = new PackageFileCreator(p.InputPath);
        }
        catch (ConstraintException e)
        {
            // Too many colors
            logger.LogError("Image file exceeded Excel's limitations.");
            logger.Log(LogLevel.None, e.Message);
            logger.LogInformation("Geez, what is that image?");
            return;
        }

        try
        {
            logger.Log(LogLevel.Information, "Writing metadata...", false);
            package.WriteMetadata(version);
            Console.WriteLine(" Done.");

            logger.LogInformation($"Temp folder is at: {package.TempDirectoryPath}");

            WriteStylesAgain: // I am sorry
            if (p.QuantizeImage)
            {
                logger.Log(LogLevel.Information, "Quantizing color...", false);
                // TODO: add image quantizing here
                Console.WriteLine(" Done.");
            }

            logger.Log(LogLevel.Information, "Writing styles...", false);
            try
            {
                package.WriteStyles();
                Console.WriteLine(" Done.");
            }
            catch (ArgumentOutOfRangeException e)
            {
                // Too many colors
                Console.WriteLine(" FAILED.");
                logger.LogError("Image file exceeded Excel's limitations.");
                logger.Log(LogLevel.None, e.Message);

                Console.Write("Try again with color quantizing? [Y/n]: ");
                string ans = Console.ReadLine()!.ToLower().Trim();
                if (ans != "y" && ans != "")
                {
                    logger.LogInformation("Operation aborted. Have a nice day owo");
                    return;
                }

                p.QuantizeImage = true;
                goto WriteStylesAgain;
            }

            logger.Log(LogLevel.Information, "Writing sheet...", false);
            package.WriteSheet();
            Console.WriteLine(" Done.");

            logger.Log(LogLevel.Information, "Compressing to package...", false);
            package.Save(p.OutputPath);
            Console.WriteLine(" Done.");

            logger.Log(LogLevel.Information, "Cleaning up temp folder...", false);
            package.Dispose();
            Console.WriteLine(" Done.");

            logger.LogInformation("All done, my master! owo");
        }
        catch (Exception e)
        {
            Console.WriteLine(" FAILED.");
            logger.LogError("Oopsie! An unknown exception has been thrown.");
            logger.LogError("We don't know what went wrong yet, but here's the details:");
            logger.Log(LogLevel.None, e.Message);
            Console.WriteLine("My dearest apologies, master~~"); // man this is pure cringe
        }
    }
}