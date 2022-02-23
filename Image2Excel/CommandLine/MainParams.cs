using Cocona;
using Image2Excel.Core;

namespace Image2Excel.CommandLine;

public class MainParams : ICommandParameterSet
{
    [Option('q', Description = "Quantize the input image to fit Excel's 256-colors-maximum requirement")]
    public bool QuantizeImage { get; set; }
    
    [Option('s', Description = "Stay silent and just do the work")]
    public bool Silent { get; set; }
    [Option('f', Description = "Force overwrite output file if already existed. Be careful with this.")]
    public bool Force { get; set; }

    [Argument(Description = "Input image file's path")]
    public string InputPath { get; set; } = "";
    
    [HasDefaultValue]
    [Argument(Description = "Output Excel file path. If left out will default to input path with \".xlsx\" appended")]
    public string? OutputPath { get; set; }

    [HasDefaultValue]
    [Option(Description = "Quantize method to use")]
    public QuantizeMethod QuantizeMethod { get; set; } = QuantizeMethod.Wu;

    [HasDefaultValue]
    [Option(Description = "Adjust the amount of dither (0 to 1)")]
    public float DitherScale { get; set; } = 0.75f;
    
    [HasDefaultValue]
    [Option(Description = "Set the dithering method used to quantize image")]
    public DitherMethod DitherMethod { get; set; } = DitherMethod.FloydSteinberg;
}