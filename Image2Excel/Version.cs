using Image2Excel.Core.Internal;

namespace Image2Excel;

public class Version : IVersion
{
    private const bool s_isPreRelease = true;
    private const int s_major = 0;
    private const int s_minor = 0;
    private static readonly string s_versionString = $"v{s_major}.{s_minor}{(s_isPreRelease ? "pre" : "")}";
    public bool IsPreRelease => s_isPreRelease;
    public int Major => s_major;
    public int Minor => s_minor;
    public string VersionString => s_versionString;

}