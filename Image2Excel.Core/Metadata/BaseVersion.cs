namespace Image2Excel.Core.Metadata;

public abstract class BaseVersion
{
    public abstract bool IsPreRelease { get; }
    public abstract int Major { get; }
    public abstract int Minor { get; }
    public string VersionString => $"v{Major}.{Minor}{(IsPreRelease ? "pre" : "")}";
}