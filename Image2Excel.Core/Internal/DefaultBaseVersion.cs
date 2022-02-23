using Image2Excel.Core.Metadata;

namespace Image2Excel.Core.Internal;

internal class DefaultVersion : BaseVersion
{
    public override bool IsPreRelease => false;
    public override int Major => 0;
    public override int Minor => 1;
}