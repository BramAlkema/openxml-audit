using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

static OpenXmlPackage OpenPackage(string path)
{
    var ext = Path.GetExtension(path).ToLowerInvariant();
    return ext switch
    {
        ".pptx" or ".pptm" or ".potx" or ".potm" => PresentationDocument.Open(path, false),
        ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => SpreadsheetDocument.Open(path, false),
        ".docx" or ".docm" or ".dotx" or ".dotm" => WordprocessingDocument.Open(path, false),
        _ => throw new ArgumentException($"Unsupported file extension '{ext}'")
    };
}

if (args.Length == 0)
{
    Console.Error.WriteLine("Usage: dotnet run -- <path-to-ooxml>");
    return;
}

var path = args[0];
if (!File.Exists(path))
{
    Console.Error.WriteLine($"File not found: {path}");
    return;
}

using var package = OpenPackage(path);
var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
var errors = validator.Validate(package);

int count = 0;
foreach (var error in errors)
{
    count++;
    Console.WriteLine($"{count}. {error.Description}");
    Console.WriteLine($"   Part: {error.Part?.Uri}");
    Console.WriteLine($"   Path: {error.Path?.XPath}");
    Console.WriteLine($"   Node: {error.Node?.LocalName}");
}

Console.WriteLine($"Errors: {count}");
