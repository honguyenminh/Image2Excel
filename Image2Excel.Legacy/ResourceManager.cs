using System.Reflection;
using System.Resources;

namespace Image2Excel.Legacy;

internal static class ResourceManager
{
    private const string ResPathPrefix = "Image2Excel.Legacy.Resource.";
    private static readonly Assembly s_assembly = typeof(Program).Assembly;

    /// <summary>
    /// Get the stream to an embedded resource inside the assembly
    /// </summary>
    /// <param name="resourcePath">Path to resource (i.e. OuterFolder.InnerFolder.FileName)</param>
    /// <returns>Stream to the resource found</returns>
    /// <exception cref="MissingManifestResourceException">Cannot find resource with given path</exception>
    public static Stream GetResourceStream(string resourcePath)
    {
        return s_assembly.GetManifestResourceStream(ResPathPrefix + resourcePath)
            ?? throw new MissingManifestResourceException("Cannot find resource " + resourcePath);
    }

    /// <summary>
    /// Get the content of an embedded resource inside the assembly as a string
    /// </summary>
    /// <param name="resourcePath">Path to resource (i.e. OuterFolder.InnerFolder.FileName)</param>
    /// <returns>The content of the resource found as a string</returns>
    /// <exception cref="MissingManifestResourceException">Cannot find resource with given path</exception>
    public static string GetResourceContent(string resourcePath)
    {
        using var stream = GetResourceStream(resourcePath);
        using StreamReader source = new(stream);
        return source.ReadToEnd();
    }

    public static void WriteResourceToFile(string resourcePath, string filePath)
    {
        using var resourceStream = GetResourceStream(resourcePath);
        using var fileStream = File.OpenWrite(filePath);
        resourceStream.CopyTo(fileStream);
    }
    public static void WriteResourceToStream(string resourcePath, Stream stream)
    {
        using var resourceStream = GetResourceStream(resourcePath);
        resourceStream.CopyTo(stream);
    }
}