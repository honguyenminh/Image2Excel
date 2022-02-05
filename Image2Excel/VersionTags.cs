namespace Image2Excel;

public static class VersionTags
{
    public const bool IsPreRelease = true;
    public const int Major = 0;
    public const int Minor = 1;
    public static readonly string Version = $"v{Major}.{Minor}" + (IsPreRelease ? "pre" : "");
}