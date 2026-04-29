// Open XML SDK runtime validator runner.
//
// Walks a corpus root, opens each OOXML file with the appropriate
// OpenXmlPackage subclass, runs OpenXmlValidator at every requested
// FileFormatVersion, and emits per-file JSON results.
//
// The output is consumed by a sibling comparator that diffs against
// the Python validator's snapshot to produce the SDK-runtime parity
// signal (Spec 013 OQ8 option B). Output schema is documented in
// README.md.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OpenXmlAudit.Parity;

internal static class Program
{
    private static readonly Dictionary<string, FileFormatVersions> KnownVersions = new()
    {
        ["Office2007"] = FileFormatVersions.Office2007,
        ["Office2010"] = FileFormatVersions.Office2010,
        ["Office2013"] = FileFormatVersions.Office2013,
        ["Office2016"] = FileFormatVersions.Office2016,
        ["Office2019"] = FileFormatVersions.Office2019,
        ["Microsoft365"] = FileFormatVersions.Microsoft365,
    };

    private static readonly string[] DefaultVersions =
    {
        "Office2007", "Office2010", "Office2013", "Office2016", "Microsoft365",
    };

    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docx", ".docm", ".dotx", ".dotm",
        ".xlsx", ".xlsm", ".xltx", ".xltm",
        ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm",
    };

    private static int Main(string[] args)
    {
        var parsed = ParseArgs(args);
        if (parsed == null)
        {
            PrintUsage();
            return 2;
        }

        if (!Directory.Exists(parsed.InputRoot))
        {
            Console.Error.WriteLine($"input root does not exist: {parsed.InputRoot}");
            return 2;
        }

        var versions = parsed.Versions.Select(v => (Name: v, Value: KnownVersions[v])).ToArray();

        var files = Directory
            .EnumerateFiles(parsed.InputRoot, "*", SearchOption.AllDirectories)
            .Where(p => SupportedExtensions.Contains(Path.GetExtension(p)))
            .OrderBy(p => p, StringComparer.Ordinal)
            .ToArray();

        var fileResults = new List<FileResult>(files.Length);
        foreach (var path in files)
        {
            var relpath = Path.GetRelativePath(parsed.InputRoot, path).Replace(Path.DirectorySeparatorChar, '/');
            var size = new FileInfo(path).Length;
            var validations = new List<VersionValidation>(versions.Length);

            foreach (var (name, value) in versions)
            {
                var validation = ValidateFile(path, value);
                validations.Add(new VersionValidation(name, validation.ErrorCount, validation.Errors, validation.OpenError));
            }

            fileResults.Add(new FileResult(relpath, size, validations));
        }

        var snapshot = new RuntimeSnapshot(
            GeneratedAtUtc: DateTime.UtcNow.ToString("o"),
            SdkPackageVersion: typeof(OpenXmlPackage).Assembly.GetName().Version?.ToString() ?? "unknown",
            ValidatorVersions: parsed.Versions,
            InputRoot: Path.GetFullPath(parsed.InputRoot),
            FileCount: fileResults.Count,
            Files: fileResults
        );

        var json = JsonSerializer.Serialize(snapshot, new JsonSerializerOptions
        {
            WriteIndented = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = SnakeCaseNamingPolicy.Instance,
        });

        if (parsed.OutputPath is { } outputPath)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath))!);
            File.WriteAllText(outputPath, json);
            Console.Error.WriteLine($"wrote {fileResults.Count} files × {versions.Length} versions → {outputPath}");
        }
        else
        {
            Console.Out.Write(json);
        }

        return 0;
    }

    private static SingleValidation ValidateFile(string path, FileFormatVersions version)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        try
        {
            return ext switch
            {
                ".docx" or ".docm" or ".dotx" or ".dotm" => RunValidator(WordprocessingDocument.Open(path, false), version),
                ".xlsx" or ".xlsm" or ".xltx" or ".xltm" => RunValidator(SpreadsheetDocument.Open(path, false), version),
                ".pptx" or ".pptm" or ".potx" or ".potm" or ".ppsx" or ".ppsm" => RunValidator(PresentationDocument.Open(path, false), version),
                _ => new SingleValidation(0, new List<ValidationErrorRecord>(), $"unsupported extension: {ext}"),
            };
        }
        catch (Exception ex)
        {
            return new SingleValidation(0, new List<ValidationErrorRecord>(), $"open failed: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static SingleValidation RunValidator(OpenXmlPackage package, FileFormatVersions version)
    {
        using (package)
        {
            var validator = new OpenXmlValidator(version);
            var errors = validator.Validate(package);
            var records = new List<ValidationErrorRecord>();
            var count = 0;
            foreach (var err in errors)
            {
                count++;
                records.Add(new ValidationErrorRecord(
                    Id: err.Id ?? string.Empty,
                    ErrorType: err.ErrorType.ToString(),
                    Part: err.Part?.Uri.ToString() ?? string.Empty,
                    Path: err.Path?.XPath ?? string.Empty,
                    Description: err.Description ?? string.Empty
                ));
            }
            return new SingleValidation(count, records, null);
        }
    }

    private static ParsedArgs? ParseArgs(string[] args)
    {
        string? inputRoot = null;
        string? outputPath = null;
        var versions = new List<string>();
        for (var i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--input-root" or "--files-root":
                    inputRoot = args[++i];
                    break;
                case "--output":
                    outputPath = args[++i];
                    break;
                case "--version":
                    versions.Add(args[++i]);
                    break;
                case "-h" or "--help":
                    return null;
                default:
                    Console.Error.WriteLine($"unknown arg: {args[i]}");
                    return null;
            }
        }
        if (inputRoot == null)
        {
            Console.Error.WriteLine("--input-root is required");
            return null;
        }
        if (versions.Count == 0)
        {
            versions.AddRange(DefaultVersions);
        }
        foreach (var v in versions)
        {
            if (!KnownVersions.ContainsKey(v))
            {
                Console.Error.WriteLine($"unknown version '{v}'. known: {string.Join(", ", KnownVersions.Keys)}");
                return null;
            }
        }
        return new ParsedArgs(inputRoot, outputPath, versions.ToArray());
    }

    private static void PrintUsage()
    {
        Console.Error.WriteLine("Usage: openxml-validator-runner --input-root <path> [--output <json>] [--version <name> ...]");
        Console.Error.WriteLine($"  Default versions: {string.Join(", ", DefaultVersions)}");
        Console.Error.WriteLine($"  Known versions:   {string.Join(", ", KnownVersions.Keys)}");
    }

    private record ParsedArgs(string InputRoot, string? OutputPath, string[] Versions);

    private record SingleValidation(int ErrorCount, List<ValidationErrorRecord> Errors, string? OpenError);
}

internal record ValidationErrorRecord(
    string Id,
    string ErrorType,
    string Part,
    string Path,
    string Description);

internal record VersionValidation(
    string Version,
    int ErrorCount,
    List<ValidationErrorRecord> Errors,
    string? OpenError);

internal record FileResult(
    string SourceRelpath,
    long SizeBytes,
    List<VersionValidation> Validations);

internal record RuntimeSnapshot(
    string GeneratedAtUtc,
    string SdkPackageVersion,
    string[] ValidatorVersions,
    string InputRoot,
    int FileCount,
    List<FileResult> Files);

internal sealed class SnakeCaseNamingPolicy : JsonNamingPolicy
{
    public static readonly SnakeCaseNamingPolicy Instance = new();

    public override string ConvertName(string name)
    {
        if (string.IsNullOrEmpty(name))
        {
            return name;
        }
        var sb = new System.Text.StringBuilder(name.Length + 4);
        for (var i = 0; i < name.Length; i++)
        {
            var c = name[i];
            if (char.IsUpper(c))
            {
                if (i > 0 && !char.IsUpper(name[i - 1]))
                {
                    sb.Append('_');
                }
                sb.Append(char.ToLowerInvariant(c));
            }
            else
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }
}
