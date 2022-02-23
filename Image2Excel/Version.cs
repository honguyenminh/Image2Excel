using Image2Excel.Core.Internal;
using Image2Excel.Core.Metadata;

namespace Image2Excel;

public class Version : BaseVersion
{
    private const bool s_isPreRelease = true;
    private const int s_major = 0;
    private const int s_minor = 0;
    public override bool IsPreRelease => s_isPreRelease;
    public override int Major => s_major;
    public override int Minor => s_minor;
}