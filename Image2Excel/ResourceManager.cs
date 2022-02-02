using System.Reflection;
using System.Resources;

namespace Image2Excel;

internal class ResourceManager
{
    private const string ResPathPrefix = "Image2Excel.Resource.";
    private readonly Assembly _assembly;

    public ResourceManager()
    {
        _assembly = Assembly.GetExecutingAssembly();
    }

    public Stream GetResourceStream(string resourcePath)
    {
        return _assembly.GetManifestResourceStream(ResPathPrefix + resourcePath)
            ?? throw new MissingManifestResourceException("Cannot find resource " + resourcePath);
    }

    public string GetEmbeddedResourceContent(string resourcePath)
    {
        using var stream = _assembly.GetManifestResourceStream(ResPathPrefix + resourcePath);
        if (stream is null)
            throw new MissingManifestResourceException("Cannot find resource " + resourcePath);
        using StreamReader source = new(stream);
        string fileContent = source.ReadToEnd();
        return fileContent;
    }
}