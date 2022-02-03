using System.Text;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Image2Excel;

/// <summary>
/// DO NOT run this in parallel
/// </summary>
public class ExcelHandler : IDisposable
{
    private readonly string _tempDirectory;
    private readonly Dictionary<string, int> _cellStyleId = new();
    
    public ExcelHandler()
    {
        // Get a non-existing temporary directory name
        do
        {
            // Good luck making a duplicate folder of this.
            // I mean it, good luck. You also need to time it right owo.
            _tempDirectory = "Image2Excel_" + Guid.NewGuid() + DateTime.Now.ToString("o");
        } while (Directory.Exists(_tempDirectory));

        Directory.CreateDirectory(_tempDirectory);
        Directory.SetCurrentDirectory(_tempDirectory);
        InitDirectories();
        InitGlobalScopedFile();
        InitWorkbook();
    }

    /// <summary>
    /// Create needed directories
    /// </summary>
    private static void InitDirectories()
    {
        Directory.CreateDirectory(@"_rels");
        Directory.CreateDirectory("docProps");
        Directory.CreateDirectory(Path.Join("xl", "worksheets"));
        Directory.CreateDirectory(Path.Join("xl", @"_rels"));
    }

    /// <summary>
    /// Write necessary global-scoped metadata files
    /// </summary>
    private static void InitGlobalScopedFile()
    {
        // Write all literal metadata files (no change needed, just copy and paste)
        ResourceManager.WriteResourceToFile(@"FileInit..rels", Path.Join(@"_rels", @".rels"));
        ResourceManager.WriteResourceToFile(@"FileInit.[Content_Types].xml", "[Content_Types].xml");
        
        // Write metadata files with replaced content
        string appMetadata = ResourceManager.GetResourceContent(@"FileInit.app.xml");
        appMetadata = appMetadata.Replace("|Ver|", VersionTags.Version);
        File.WriteAllText(Path.Join("docProps", "app.xml"), appMetadata);
        
        StringBuilder coreMetadata = new(ResourceManager.GetResourceContent("FileInit.core.xml"));
        coreMetadata.Replace("|CurrentUsername|", Environment.UserName);
        // Format current time as W3CDTF format
        string formattedCurrentTime = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssK");
        coreMetadata.Replace("|CreationTime|", formattedCurrentTime);
        File.WriteAllText(Path.Join("docProps", "core.xml"), coreMetadata.ToString());
    }

    private static void InitWorkbook()
    {
        Directory.SetCurrentDirectory("xl");
        ResourceManager.WriteResourceToFile("WorkbookInit.workbook.xml", "workbook.xml");
        ResourceManager.WriteResourceToFile("WorkbookInit.workbook.xml.rels", Path.Join("_rels", "workbook.xml.rels"));
    }

    // TODO: replace dimension in sheet1 head
    public void WriteStyles(Image<Rgb24> image)
    {
        // Process image
        StringBuilder fillStyles = new();
        StringBuilder cellStyles = new();
        for (int y = 0; y < image.Height; y++)
        {
            if (_cellStyleId.Count > 255)
            {
                Console.WriteLine("Too many colors. Max is 256 (blame Excel owo)");
                return;
            }

            Span<Rgb24> row = image.GetPixelRowSpan(y);
            for (int x = 0; x < image.Width; x++)
            {
                string hexCode = row[x].ToHex();
                if (_cellStyleId.ContainsKey(hexCode)) continue;

                fillStyles.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"");
                fillStyles.Append(hexCode);
                fillStyles.Append("\"/><bgColor indexed=\"64\"/></patternFill></fill>");

                _cellStyleId[hexCode] = _cellStyleId.Count + 1;
                cellStyles.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"");
                cellStyles.Append(_cellStyleId.Count + 1);
                cellStyles.Append("\" borderId=\"0\" xfId=\"0\" applyFill=\"1\"/>");
            }
        }
    }

    private void ReleaseUnmanagedResources()
    {
        if (Directory.Exists(_tempDirectory))
            Directory.Delete(_tempDirectory, true);
    }

    public void Dispose()
    {
        ReleaseUnmanagedResources();
        GC.SuppressFinalize(this);
    }
    ~ExcelHandler()
    {
        ReleaseUnmanagedResources();
    }
}