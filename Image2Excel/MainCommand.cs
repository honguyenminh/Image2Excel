using System.Data;
using Image2Excel.CommandLine;
using Image2Excel.Core;
using Microsoft.Extensions.Logging;
using Version = Image2Excel.Core.Metadata.Version;

namespace Image2Excel;

internal class MainCommand
{
    private readonly ConsoleLogger _logger;
    private readonly Version _version;
    public MainCommand(ConsoleLogger logger, Version version)
    {
        _logger = logger;
        _version = version;
    }

    public void Command(MainParams p)
    {
        _logger.LowestLogLevel = p.Silent ? LogLevel.Warning : LogLevel.Information;
        _logger.LogInformation($"Image2Excel CLI {_version.VersionString}");
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
                _logger.LogError($"Error message: {e.Message}", logHead: false);
            }
        }
        
        try
        {
            PackageFileCreator package;
            try
            {
                package = new PackageFileCreator(p.InputPath);
            }
            catch (ConstraintException e)
            {
                // Too many colors
                _logger.LogError("Image file exceeded Excel's limitations.");
                _logger.LogError(e.Message, emptyHead: true);
                _logger.LogInformation("Geez, what is that image?");
                return;
            }

            _logger.LogInformation("Writing metadata...", false);
            package.WriteMetadata(_version);
            _logger.LogInformation(" Done.", logHead: false);

            _logger.LogInformation($"Temp folder is at: {package.TempDirectoryPath}");

            WriteStylesAgain:
            if (p.QuantizeImage)
            {
                _logger.LogInformation("Quantizing color...", false);
                package.QuantizeImage(p.QuantizeMethod, p.DitherMethod, p.DitherScale);
                _logger.LogInformation(" Done.", logHead: false);
            }

            _logger.LogInformation("Writing styles...", false);
            try
            {
                package.WriteStyles();
                _logger.LogInformation(" Done.", logHead: false);
            }
            catch (ArgumentOutOfRangeException e)
            {
                // Too many colors
                _logger.LogInformation(" FAILED.", logHead: false);
                _logger.LogError("Failed writing styles. Image file exceeded Excel's limitations.");
                _logger.LogError(e.Message);

                Console.Write("******  Try again with color quantizing? [Y/n]: ");
                string ans = Console.ReadLine()!.ToLower().Trim();
                if (ans != "y" && ans != "")
                {
                    _logger.LogInformation("Operation aborted. Have a nice day owo");
                    return;
                }

                p.QuantizeImage = true;
                goto WriteStylesAgain;
            }

            _logger.LogInformation("Writing sheet...", false);
            package.WriteSheet();
            _logger.LogInformation(" Done.", logHead: false);

            _logger.LogInformation("Compressing to package...", false);
            package.Save(p.OutputPath);
            _logger.LogInformation(" Done.", logHead: false);

            _logger.LogInformation("Cleaning up temp folder...", false);
            package.Dispose();
            _logger.LogInformation(" Done.", logHead: false);

            _logger.LogInformation("All done, my master! owo");
        }
        catch (Exception e)
        {
            _logger.LogInformation(" FAILED.", logHead: false);
            _logger.LogError("Oopsie! An unknown exception has been thrown.");
            _logger.LogError("We don't know what went wrong yet, but here's the details:", emptyHead: true);
            _logger.LogError(e.Message);
            _logger.LogError("My dearest apologies, master~~"); // man this is pure cringe
        }
    }
}