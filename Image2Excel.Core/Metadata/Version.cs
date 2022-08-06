namespace Image2Excel.Core.Metadata;
// Keep it simple, stupid
// TODO: add xmldoc
public record Version
(
    int Major,
    int Minor,
    bool IsPreRelease
)
{
    public string VersionString { get; } = $"v{Major}.{Minor}{(IsPreRelease ? "pre" : "")}";

    public static Version Default { get; } = new Version(0, 1, false);
}