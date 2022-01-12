namespace Image2Excel;

public static class VersionTags
{
    public const bool IsPreproduction = true;
    public const int MajorVersion = 0;
    public const int MinorVersion = 1;
    public static readonly string Version = $"v{MajorVersion}.{MinorVersion}" + (IsPreproduction ? "pre" : "");
}