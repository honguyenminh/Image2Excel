using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Processing.Processors.Quantization;

namespace Image2Excel.Core.Internal;

public static class Extensions
{
    public static string ToArgbHex(this Rgba32 color)
    {
        uint hexOrder = (uint)(color.B << 0 | color.G << 8 | color.R << 16 | color.A << 24);
        return hexOrder.ToString("X8");
    }
}