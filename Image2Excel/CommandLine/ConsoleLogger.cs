using System.Buffers;
using Microsoft.Extensions.Logging;
using SixLabors.ImageSharp;

namespace Image2Excel.CommandLine;

public class ConsoleLogger
{
    public LogLevel LowestLogLevel { get; set; } = LogLevel.Information;

    #region Auxiliary methods invoking main log method
    public void LogTrace(string message)
    {
        Log(LogLevel.Trace, message);
    }
    public void LogDebug(string message)
    {
        Log(LogLevel.Debug, message);
    }
    public void LogInformation(string message)
    {
        Log(LogLevel.Information, message);
    }
    public void LogWarning(string message)
    {
        Log(LogLevel.Warning, message);
    }
    public void LogError(string message)
    {
        Log(LogLevel.Error, message);

    }
    public void LogCritical(string message)
    {
        Log(LogLevel.Critical, message);
    }
    #endregion

    /// <summary>
    /// Log a message with message header indicating the given log level, with c o l o r s
    /// </summary>
    /// <remarks>
    /// This WILL PRINT none level logs as usual, just with the head spacing appended
    /// </remarks>
    /// <param name="logLevel">Log level of message</param>
    /// <param name="message">Message to log</param>
    /// <param name="newLine">Append new line after message</param>
    /// <exception cref="ArgumentOutOfRangeException">Log level enum is out of supported range</exception>
    public void Log(LogLevel logLevel, string message, bool newLine = true)
    {
        if (logLevel < LowestLogLevel) return;
        // ReSharper disable once SwitchExpressionHandlesSomeKnownEnumValuesWithExceptionInDefault
        // Log heads indicating log level
        string head = logLevel switch
        {
            LogLevel.Trace       => "[Trace] ",
            LogLevel.Debug       => "[Debug] ",
            LogLevel.Information => "[Info]  ",
            LogLevel.Warning     => "[WARN]  ",
            LogLevel.Error       => "[ERROR] ",
            LogLevel.Critical    => "[CRIT]  ",
            LogLevel.None        => "        ",
            _ => throw new ArgumentOutOfRangeException(nameof(logLevel), logLevel, 
                "Log level enum is out of supported range")
        };
        // Set the  c o l o r s
        if (s_fgColor.ContainsKey(logLevel)) 
            Console.ForegroundColor = s_fgColor[logLevel];
        if (s_bgColor.ContainsKey(logLevel)) 
            Console.BackgroundColor= s_bgColor[logLevel];

        Console.Write(head);

        static void WriteSpan(ReadOnlySpan<char> span)
        {
            foreach (char c in span) Console.Write(c);
        }
        if (Console.WindowWidth < head.Length + message.Length)
        {
            ReadOnlySpan<char> span = message;
            int lineLength = Console.WindowWidth - head.Length;
            WriteSpan(span[..(lineLength - 1)]);
            Console.WriteLine();
            int i = lineLength;
            while (i < message.Length)
            {
                // Space out at the top
                (int _, int top) = Console.GetCursorPosition();
                Console.SetCursorPosition(head.Length, top);
                // last line, just print till the end
                if (message.Length - i  >= message.Length % lineLength)
                {
                    WriteSpan(span[i..]);
                    if (newLine) Console.WriteLine();
                    break;
                }
                int endPos = i + lineLength - 1;
                WriteSpan(span[i..endPos]);
                Console.WriteLine();
                i += lineLength;
            }
        }
        else Console.WriteLine(message);

        Console.ResetColor();
    }

    // Color dictionaries
    private static readonly Dictionary<LogLevel, ConsoleColor> s_fgColor = new()
    {
        {LogLevel.Trace, ConsoleColor.DarkGray},
        {LogLevel.Debug, ConsoleColor.Gray},
        {LogLevel.Warning, ConsoleColor.Yellow},
        {LogLevel.Error, ConsoleColor.Red},
        {LogLevel.Critical, ConsoleColor.DarkRed}
    };
    private static readonly Dictionary<LogLevel, ConsoleColor> s_bgColor = new()
    {
        {LogLevel.Debug, ConsoleColor.DarkGray},
        {LogLevel.Critical, ConsoleColor.DarkGray}
    };
}