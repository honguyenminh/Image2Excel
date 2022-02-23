using Image2Excel.Core.Internal;
using Image2Excel.Core.Metadata;

namespace Image2Excel.Legacy;

public class Version : BaseVersion
{
    private const bool _isPreRelease = true;
    private const int _major = 0;
    private const int _minor = 2;
    public override bool IsPreRelease => _isPreRelease;
    public override int Major => _major;
    public override int Minor => _minor;
}