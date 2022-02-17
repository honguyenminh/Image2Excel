namespace Image2Excel.Legacy;

public static class VersionTags
{
    public const bool IsPreRelease = true;
    public const int Major = 0;
    public const int Minor = 2;
    public static readonly string Version = $"v{Major}.{Minor}" + (IsPreRelease ? "pre" : "");
}