using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

static OpenXmlPackage OpenPackage(string path)
{
    var ext = Path.GetExtension(path).ToLowerInvariant();
    return ext switch
    {
        ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" or ".ppam" or ".thmx" =>
            PresentationDocument.Open(path, false),
        ".xlsx" or ".xlsm" or ".xltx" or ".xltm" or ".xlam" =>
            SpreadsheetDocument.Open(path, false),
        ".docx" or ".docm" or ".dotx" or ".dotm" =>
            WordprocessingDocument.Open(path, false),
        _ => throw new ArgumentException($"Unsupported file extension '{ext}'")
    };
}

if (args.Length < 1)
{
    Console.Error.WriteLine("Usage: dotnet run --project scripts/sdk_compare -- <path-to-ooxml> [<path> ...]");
    return 2;
}

var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
var results = new List<object>();

foreach (var path in args)
{
    if (!File.Exists(path))
    {
        results.Add(new
        {
            File = path,
            Errors = new[]
            {
                new
                {
                    Id = "",
                    Description = "File not found",
                    Part = "",
                    Path = "",
                    Node = "",
                    RelatedNode = "",
                    ErrorType = "FileNotFound",
                }
            }
        });
        continue;
    }

    try
    {
        using var package = OpenPackage(path);
        var errors = validator.Validate(package).Select(error => new
        {
            Id = error.Id,
            Description = error.Description,
            Part = error.Part?.Uri?.ToString() ?? "",
            Path = error.Path?.XPath ?? "",
            Node = error.Node?.LocalName ?? "",
            RelatedNode = error.RelatedNode?.LocalName ?? "",
            ErrorType = error.ErrorType.ToString(),
        });
        results.Add(new
        {
            File = path,
            Errors = errors.ToList(),
        });
    }
    catch (Exception ex)
    {
        results.Add(new
        {
            File = path,
            Errors = new[]
            {
                new
                {
                    Id = "",
                    Description = $"Exception: {ex.Message}",
                    Part = "",
                    Path = "",
                    Node = "",
                    RelatedNode = "",
                    ErrorType = "Exception",
                }
            }
        });
    }
}

var options = new JsonSerializerOptions { WriteIndented = true };
Console.WriteLine(JsonSerializer.Serialize(results, options));
return 0;
