using Microsoft.Extensions.Logging;

namespace Image2Excel.CommandLine;

public class ConsoleLogger
{
    public LogLevel LowestLogLevel { get; set; } = LogLevel.Information;

    #region Auxiliary methods invoking main log method
    /// <inheritdoc cref="Log"/>
    public void LogTrace(string message, bool newLine = true, bool logHead = true, bool emptyHead = false)
    {
        Log(LogLevel.Trace, message, newLine, logHead, emptyHead);
    }
    /// <inheritdoc cref="Log"/>
    public void LogDebug(string message, bool newLine = true, bool logHead = true, bool emptyHead = false)
    {
        Log(LogLevel.Debug, message, newLine, logHead, emptyHead);
    }
    /// <inheritdoc cref="Log"/>
    public void LogInformation(string message, bool newLine = true, bool logHead = true, bool emptyHead = false)
    {
        Log(LogLevel.Information, message, newLine, logHead, emptyHead);
    }
    /// <inheritdoc cref="Log"/>
    public void LogWarning(string message, bool newLine = true, bool logHead = true, bool emptyHead = false)
    {
        Log(LogLevel.Warning, message, newLine, logHead, emptyHead);
    }
    /// <inheritdoc cref="Log"/>
    public void LogError(string message, bool newLine = true, bool logHead = true, bool emptyHead = false)
    {
        Log(LogLevel.Error, message, newLine, logHead, emptyHead);
    }
    /// <inheritdoc cref="Log"/>
    public void LogCritical(string message, bool newLine = true, bool logHead = true, bool emptyHead = false)
    {
        Log(LogLevel.Critical, message, newLine, logHead, emptyHead);
    }
    #endregion

    /// <summary>
    /// Log a message with message header indicating the given log level, with c o l o r s
    /// </summary>
    /// <param name="logLevel">Log level of message</param>
    /// <param name="message">Message to log</param>
    /// <param name="newLine">Append new line after message</param>
    /// <param name="logHead">Write log level head before message</param>
    /// <param name="emptyHead">Write an empty log level head as indent</param>
    /// <exception cref="ArgumentOutOfRangeException">Log level enum is out of supported range</exception>
    public void Log(LogLevel logLevel, string message, bool newLine = true, bool logHead = true, bool emptyHead = false)
    {
        if (logLevel < LowestLogLevel) return;
        if (logLevel == LogLevel.None) return;
        // Log heads indicating log level
        // ReSharper disable once SwitchExpressionHandlesSomeKnownEnumValuesWithExceptionInDefault
        string head = logLevel switch
        {
            LogLevel.Trace       => "[Trace] ",
            LogLevel.Debug       => "[Debug] ",
            LogLevel.Information => "[Info]  ",
            LogLevel.Warning     => "[WARN]  ",
            LogLevel.Error       => "[ERROR] ",
            LogLevel.Critical    => "[CRIT]  ",
            _ => throw new ArgumentOutOfRangeException(nameof(logLevel), logLevel,
                "Log level enum is out of supported range")
        };
        if (emptyHead) head = "        ";
        if (!logHead) head = "";
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
            WriteSpan(span[..lineLength]);
            int i = lineLength;
            while (i < message.Length)
            {
                // Space out at the top
                (int _, int top) = Console.GetCursorPosition();
                Console.SetCursorPosition(head.Length, top);
                // last line, just print till the end
                if (message.Length - i  <= message.Length % lineLength)
                {
                    WriteSpan(span[i..]);
                    if (message.Length - i == lineLength) break;
                    if (newLine) Console.WriteLine();
                    break;
                }
                int endPos = i + lineLength;
                WriteSpan(span[i..endPos]);
                i += lineLength;
            }
        }
        else
        {
            if (newLine) Console.WriteLine(message);
            else Console.Write(message); 
        }

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