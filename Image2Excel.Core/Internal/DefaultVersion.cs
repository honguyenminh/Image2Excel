namespace Image2Excel.Core.Internal;

internal class DefaultVersion : IVersion
{
    public bool IsPreRelease => false;
    public int Major => 0;
    public int Minor => 1;
    public string VersionString => "v0.1";
}