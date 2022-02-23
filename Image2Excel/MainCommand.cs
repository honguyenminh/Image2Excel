using System.Data;
using Image2Excel.CommandLine;
using Image2Excel.Core;
using Microsoft.Extensions.Logging;

namespace Image2Excel;

internal class MainCommand
{
    private readonly Version _version;
    private readonly ConsoleLogger _logger;
    public MainCommand(Version version, ConsoleLogger logger)
    {
        _version = version;
        _logger = logger;
    }
    public void Command(MainParams p)
    {
        _logger.LowestLogLevel = p.Silent ? LogLevel.Warning : LogLevel.Information;
        _logger.LogInformation($"Image2Excel {_version.VersionString}");
#if DEBUG
        Console.WriteLine("--------------------DEBUG BUILD--------------------");
#endif
        _logger.LogInformation("GitHub: https://github.com/honguyenminh/Image2Excel");
        if (_version.IsPreRelease)
        {
            _logger.LogWarning("THIS IS A PRE-RELEASE, FEATURES MIGHT BE UNSTABLE!");
            Console.WriteLine("you have been warned uwu");
        }
        if (!p.Silent) Console.WriteLine("---------------------------------------------------");

        p.OutputPath ??= p.InputPath + ".xlsx"; // if output path not specified
        if (!File.Exists(p.InputPath))
        {
            _logger.LogError("Cannot find/read image file at provided path");
            return;
        }

        if (File.Exists(p.OutputPath))
        {
            if (!p.Force)
            {
                _logger.LogError($"Output file already existed at {p.OutputPath}");
                Console.Write("******  Do you want to overwrite output file? [y/N]: ");
                switch (Console.ReadLine()?.ToLower().Trim())
                {
                    case "y":
                        _logger.LogWarning("Don't blame me then owo. Overwriting file...");
                        break;
                    // Default is no
                    default:
                        _logger.LogInformation("Aborted operation. Have a nice day~");
                        return;
                }
            }
            try
            {
                File.Delete(p.OutputPath);
            }
            catch (UnauthorizedAccessException e)
            {
                _logger.LogError("Can't overwrite output file, not enough permission or is read-only");
                _logger.Log(LogLevel.None, $"Error message: {e.Message}");
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
            _logger.LogError("Image file exceeded Excel's limitations.");
            _logger.Log(LogLevel.None, e.Message);
            _logger.LogInformation("Geez, what is that image?");
            return;
        }

        try
        {
            _logger.Log(LogLevel.Information, "Writing metadata...", false);
            package.WriteMetadata(_version);
            Console.WriteLine(" Done.");

            _logger.LogInformation($"Temp folder is at: {package.TempDirectoryPath}");

            WriteStylesAgain: // I am sorry
            if (p.QuantizeImage)
            {
                _logger.Log(LogLevel.Information, "Quantizing color...", false);
                package.QuantizeImage(p.QuantizeMethod, p.DitherMethod, p.DitherScale);
                Console.WriteLine(" Done.");
            }

            _logger.Log(LogLevel.Information, "Writing styles...", false);
            try
            {
                package.WriteStyles();
                Console.WriteLine(" Done.");
            }
            catch (ArgumentOutOfRangeException e)
            {
                // Too many colors
                Console.WriteLine(" FAILED.");
                _logger.LogError("Image file exceeded Excel's limitations.");
                _logger.Log(LogLevel.None, e.Message);

                Console.Write("Try again with color quantizing? [Y/n]: ");
                string ans = Console.ReadLine()!.ToLower().Trim();
                if (ans != "y" && ans != "")
                {
                    _logger.LogInformation("Operation aborted. Have a nice day owo");
                    return;
                }

                p.QuantizeImage = true;
                goto WriteStylesAgain;
            }

            _logger.Log(LogLevel.Information, "Writing sheet...", false);
            package.WriteSheet();
            Console.WriteLine(" Done.");

            _logger.Log(LogLevel.Information, "Compressing to package...", false);
            package.Save(p.OutputPath);
            Console.WriteLine(" Done.");

            _logger.Log(LogLevel.Information, "Cleaning up temp folder...", false);
            package.Dispose();
            Console.WriteLine(" Done.");

            _logger.LogInformation("All done, my master! owo");
        }
        catch (Exception e)
        {
            Console.WriteLine(" FAILED.");
            _logger.LogError("Oopsie! An unknown exception has been thrown.");
            _logger.LogError("We don't know what went wrong yet, but here's the details:");
            _logger.Log(LogLevel.None, e.Message);
            Console.WriteLine("My dearest apologies, master~~"); // man this is pure cringe
        }
    }
}