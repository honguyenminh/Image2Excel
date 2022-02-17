using Cocona;
using Microsoft.Extensions.Logging;
using Microsoft.VisualBasic.FileIO;

namespace Image2Excel.CommandLine;

public class MainCommand
{
    public void Command(MainParams p, ILogger logger)
    {
        if (!p.Silent)
        {
            Console.WriteLine("Image2Excel.Legacy " + Version.VersionString);
            Console.WriteLine("GitHub: https://github.com/honguyenminh/Image2Excel.Legacy");
            if (Version.IsPreRelease)
            {
                Console.WriteLine("THIS IS A PRE-RELEASE, FEATURES MIGHT BE UNSTABLE!");
                Console.WriteLine("you have been warned uwu");
            }
            Console.WriteLine("---------------------------------------------------");
        }
        p.OutputPath ??= p.InputPath + ".xlsx"; // if output path not specified
        if (!File.Exists(p.InputPath))
        {
            logger.LogError("Cannot find/read image file at provided path");
            return;
        }

        if (File.Exists(p.OutputPath))
        {
            if (p.Force) goto Delete;
            logger.LogError("Output file already existed at {p.OutputPath}", p.OutputPath);
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
            Delete:
            try
            {
                File.Delete(p.OutputPath);
            }
            catch (UnauthorizedAccessException e)
            {
                logger.LogCritical(e, "Can't overwrite output file, not enough permission or is read-only");
            }
        }
        
        // TODO: handle excel here
    }

    [Hidden]
    [Command("owo")]
    public void SuperCommand()
    {
        Console.WriteLine("THIS APP HAS SUPER OWO POWER");
    }
}