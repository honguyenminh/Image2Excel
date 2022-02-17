using System.Text;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using System.IO.Compression;

namespace Image2Excel.Legacy;

public class ExcelHandler : IDisposable
{
    public string TempDirectoryPath { get; }
    private readonly Dictionary<string, int> _cellStyleId = new();
    
    public ExcelHandler()
    {
        // Get a non-existing temporary directory name
        do
        {
            // Good luck making a duplicate folder of this.
            // I mean it, good luck. You also need to time it right owo.
            TempDirectoryPath = $"Image2Excel_{Guid.NewGuid()}_{DateTime.Now.ToString("d_HH-mm-ss-fff")}";
        } while (Directory.Exists(TempDirectoryPath));

        var dirInfo = Directory.CreateDirectory(TempDirectoryPath);
        TempDirectoryPath = dirInfo.FullName;
        InitDirectories();
        InitGlobalScopedFile();
        InitWorkbook();
    }

    /// <summary>
    /// Create needed directories
    /// </summary>
    private void InitDirectories()
    {
        Directory.SetCurrentDirectory(TempDirectoryPath);
        Directory.CreateDirectory(@"_rels");
        Directory.CreateDirectory("docProps");
        Directory.CreateDirectory(Path.Join("xl", "worksheets"));
        Directory.CreateDirectory(Path.Join("xl", @"_rels"));
        Directory.SetCurrentDirectory("..");
    }

    /// <summary>
    /// Write necessary global-scoped metadata files
    /// </summary>
    private void InitGlobalScopedFile()
    {
        // Write all literal metadata files (no change needed, just copy and paste)
        ResourceManager.WriteResourceToFile("FileInit..rels", 
            Path.Join(TempDirectoryPath, "_rels", ".rels"));
        ResourceManager.WriteResourceToFile("FileInit.[Content_Types].xml", 
            Path.Join(TempDirectoryPath, "[Content_Types].xml"));

        // Write metadata files with replaced content
        string appMetadata = ResourceManager.GetResourceContent("FileInit.app.xml");
        appMetadata = appMetadata.Replace("|Ver|", $"{VersionTags.Major}.{VersionTags.Minor}");
        File.WriteAllText(Path.Join(TempDirectoryPath, "docProps", "app.xml"), appMetadata);

        StringBuilder coreMetadata = new(ResourceManager.GetResourceContent("FileInit.core.xml"));
        coreMetadata.Replace("|CurrentUsername|", Environment.UserName);
        // Format current time as W3CDTF format
        string formattedCurrentTime = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssZ");
        coreMetadata.Replace("|CreationTime|", formattedCurrentTime);
        File.WriteAllText(Path.Join(TempDirectoryPath, "docProps", "core.xml"), coreMetadata.ToString());
    }

    private void InitWorkbook()
    {
        ResourceManager.WriteResourceToFile("WorkbookInit.workbook.xml", 
            Path.Join(TempDirectoryPath, "xl", "workbook.xml"));
        ResourceManager.WriteResourceToFile("WorkbookInit.workbook.xml.rels", 
            Path.Join(TempDirectoryPath, "xl", "_rels", "workbook.xml.rels"));
    }

    // TODO: replace dimension in sheet1 head
    /// <summary>
    /// Write styles file for the given image
    /// </summary>
    /// <param name="image">Image to get color styles from</param>
    /// <exception cref="ArgumentOutOfRangeException">Too many colors in image</exception>
    public void WriteStyles(Image<Rgba32> image)
    {
        // The fill/cell styles strings to write to styles.xml
        // TODO: save these to separate temp file
        StringBuilder fillStyles = new();
        StringBuilder cellStyles = new();

        for (int y = 0; y < image.Height; y++)
        {
            // TODO: add default colors from excel here to avoid duplication
            Span<Rgba32> row = image.GetPixelRowSpan(y);
            for (int x = 0; x < image.Width; x++)
            {
                if (_cellStyleId.Count > 256)
                {
                    _cellStyleId.Clear();
                    throw new ArgumentOutOfRangeException(nameof(image), 
                        "Too many colors. Max is 256 (blame Excel owo)");
                }

                string hexCode = row[x].ToArgbHex();
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

        // Write styles to file
        WriteStylesToFile(fillStyles.ToString(), cellStyles.ToString());
    }

    private void WriteStylesToFile(string fillStyles, string cellStyles)
    {
        using var fileStream = File.OpenWrite(Path.Join(TempDirectoryPath, "xl", "styles.xml"));
        using StreamWriter writer = new(fileStream);

        ResourceManager.WriteResourceToStream("Parts.styles.xml.head.txt", fileStream);

        // Write fill styles
        writer.Write("<fills count=\"");
        writer.Write(_cellStyleId.Count + 2);
        // Two mandatory styles, blame excel
        writer.Write("\"><fill><patternFill patternType=\"none\"/></fill>"
                     + "<fill><patternFill patternType=\"gray125\"/></fill>");

        writer.Write(fillStyles);
        writer.Flush();

        ResourceManager.WriteResourceToStream("Parts.styles.xml.mid.txt", fileStream);

        // Write cell styles
        writer.Write($"<cellXfs count=\"{_cellStyleId.Count + 1}\">");
        // Dummy one, because excel is stupid
        writer.Write("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
        writer.Write(cellStyles);
        writer.Flush();

        ResourceManager.WriteResourceToStream("Parts.styles.xml.bottom.txt", fileStream);
    }

    public void WriteSheet(Image<Rgba32> image)
    {
        using var fileStream = File.OpenWrite(Path.Join(TempDirectoryPath, "xl", "worksheets", "sheet1.xml"));
        using StreamWriter writer = new(fileStream);

        ResourceManager.WriteResourceToStream("Parts.sheet1.xml.head.txt", fileStream);

        // Write sheet data to temp file
        string? highestColumnName = WriteTempSheet(image, "sheet1.xml.temp");

        if (highestColumnName is null)
        {
            writer.Write("<dimension ref=\"A1:A1\"/>"); // Image is empty
        }
        else writer.Write($"<dimension ref=\"A1:{highestColumnName}{image.Height}\"/>");
        writer.Write("<sheetFormatPr defaultColWidth=\"1\" defaultRowHeight=\"5.95\" customHeight=\"1\"/>");
        writer.Flush();

        // Read temp file back and save to main sheet file
        string tempPath = Path.Join(TempDirectoryPath, "xl", "worksheets", "sheet1.xml.temp");
        using (var tempFile = File.OpenRead(tempPath))
        {
            tempFile.CopyTo(fileStream);
        }
        File.Delete(tempPath);

        ResourceManager.WriteResourceToStream("Parts.sheet1.xml.bottom.txt", fileStream);
    }

    /// <summary>
    /// Write a temp file containing sheetData for a sheet inside the sheets folder
    /// </summary>
    /// <param name="image">Image to get data from</param>
    /// <param name="tempName">Name of the temp file</param>
    /// <returns>The highest column name of the sheet</returns>
    private string? WriteTempSheet(Image<Rgba32> image, string tempName)
    {
        using var fileStream = File.OpenWrite(Path.Join(TempDirectoryPath, "xl", "worksheets", tempName));
        using StreamWriter writer = new(fileStream);

        writer.Write("<sheetData>");
        char[] highestColChars = Array.Empty<char>();
        for (int rowIndex = 0; rowIndex < image.Height; rowIndex++)
        {
            writer.Write($"<row r=\"{rowIndex + 1}\">");
            // Build cell list
            List<char> columnName = new() { 'A' };
            Span<Rgba32> row = image.GetPixelRowSpan(rowIndex);
            for (int x = 0; x < image.Width; x++)
            {
                string hexCode = row[x].ToArgbHex();
                highestColChars = columnName.ToArray();
                writer.Write("<c r=\"");
                writer.Write(highestColChars);
                writer.Write(rowIndex + 1);
                writer.Write("\" s=\"");
                writer.Write(_cellStyleId[hexCode]);
                writer.Write("\"/>");
                
                // Increase column name
                bool keepRunning = true;
                int index = columnName.Count - 1;
                while (keepRunning && index > -1)
                {
                    if (columnName[index] == 'Z')
                    {
                        columnName[index] = 'A';
                        index--;
                        continue;
                    }
                    columnName[index]++;
                    keepRunning = false;
                }
                if (index == -1)
                {
                    columnName.Insert(0, 'A');
                }
            }

            writer.Write("</row>");
        }
        writer.Write("</sheetData>");
        writer.Flush();

        return highestColChars.Length != 0 ? new string(highestColChars) : null;
    }

    public void Save(string filePath)
    {
        if (File.Exists(filePath))
            throw new ArgumentException("File already existed");
        
        ZipFile.CreateFromDirectory(TempDirectoryPath, filePath);
    }

    private void ReleaseUnmanagedResources()
    {
        if (Directory.Exists(TempDirectoryPath))
            Directory.Delete(TempDirectoryPath, true);
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