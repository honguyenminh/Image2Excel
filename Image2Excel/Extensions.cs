using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Processing.Processors.Quantization;

namespace Image2Excel;

public static class Extensions
{
    public static string ToArgbHex(this Rgba32 color)
    {
        uint hexOrder = (uint)(color.B << 0 | color.G << 8 | color.R << 16 | color.A << 24);
        return hexOrder.ToString("X8");
    }
    
    private static readonly QuantizerOptions s_options = new()
    {
        Dither = KnownDitherings.FloydSteinberg,
        DitherScale = 0.75f
    };
    
    private static readonly WuQuantizer s_wuQuantizer = new(s_options);
    private static readonly OctreeQuantizer s_octreeQuantizer = new(s_options);
    
    /// <summary>
    /// Quantize (aka convert to indexed color) the given image to 256 color with Xiaolin Wu's quantizer
    /// </summary>
    /// <remarks>
    /// Produce high quality color
    /// </remarks>
    public static void QuantizeColorWu(this Image<Rgba32> image)
    {
        image.Mutate(context =>
        {
            context.Quantize(s_wuQuantizer);
        });
    }
    
    /// <summary>
    /// Quantize (aka convert to indexed color) the given image to 256 color with Octree quantizer
    /// </summary>
    /// <remarks>
    /// Produce okay quality, run fast
    /// </remarks>
    public static void QuantizeColorOctree(this Image<Rgba32> image)
    {
        image.Mutate(context =>
        {
            context.Quantize(s_octreeQuantizer);
        });
    }
}