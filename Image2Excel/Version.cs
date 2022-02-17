namespace Image2Excel;

public class Version
{
    public const bool IsPreRelease = true;
    public const int Major = 0;
    public const int Minor = 2;
    public static readonly string VersionString = $"v{Major}.{Minor}{(IsPreRelease ? "pre" : "")}";
}