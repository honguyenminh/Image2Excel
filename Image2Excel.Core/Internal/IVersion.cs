namespace Image2Excel.Core.Internal;

public interface IVersion
{
    public bool IsPreRelease { get; }
    public int Major { get; }
    public int Minor { get; }
    public string VersionString { get; }
}