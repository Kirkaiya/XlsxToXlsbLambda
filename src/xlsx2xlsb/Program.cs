using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using XlsxToXlsb;

namespace XlsxToXlsb.Cli;

/// <summary>
/// <c>xlsx2xlsb &lt;input.xlsx&gt; &lt;output.xlsb&gt;</c>
///
/// Converts one workbook and prints ONE line of JSON to stdout, success or failure, so a caller
/// (the Evolve agent's Python tool) parses stdout and never scrapes text:
///
///   {"ok":true,"outputPath":"…","sizeBytes":130651,"sheetCount":22,"adjustments":{…}}
///   {"ok":false,"error":"…"}
///
/// The error is also written to stderr for a human at a terminal. Exit codes:
///   0 converted · 1 the workbook cannot be converted · 2 usage · 3 the files could not be read or written.
/// The output is written to a temporary file beside the target and moved into place, so a failed
/// run never leaves a partial .xlsb behind.
/// </summary>
public static class Program
{
    public const int Ok = 0;
    public const int ConversionFailed = 1;
    public const int UsageError = 2;
    public const int IoError = 3;

    private const string Usage = "usage: xlsx2xlsb <input.xlsx> <output.xlsb>   |   xlsx2xlsb --version";

    private static readonly JsonSerializerOptions Json = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public static int Main(string[] args) => Run(args, Console.Out, Console.Error);

    /// <summary>The whole program, with its streams injected so tests can run it in-process.</summary>
    internal static int Run(string[] args, TextWriter stdout, TextWriter stderr)
    {
        if (args is ["--version"])
        {
            stdout.WriteLine(Version());
            return Ok;
        }
        if (args.Length != 2 || args.Any(string.IsNullOrWhiteSpace))
            return Fail(stdout, stderr, UsageError, Usage);

        string input = Path.GetFullPath(args[0]);
        string output = Path.GetFullPath(args[1]);
        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase))
            return Fail(stdout, stderr, UsageError, "the output path must differ from the input path");
        if (!output.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase))
            return Fail(stdout, stderr, UsageError, $"the output file must end in .xlsb: {output}");

        byte[] xlsx;
        try
        {
            xlsx = File.ReadAllBytes(input);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            return Fail(stdout, stderr, IoError, $"cannot read {input}: {e.Message}");
        }

        ConversionResult result;
        try
        {
            result = XlsbConverter.Convert(xlsx);
        }
        catch (XlsbConversionException e)
        {
            return Fail(stdout, stderr, ConversionFailed, e.Message);
        }

        string temporary = $"{output}.{Guid.NewGuid():N}.tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(output)!);
            File.WriteAllBytes(temporary, result.Xlsb);
            File.Move(temporary, output, overwrite: true);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            TryDelete(temporary);
            return Fail(stdout, stderr, IoError, $"cannot write {output}: {e.Message}");
        }

        stdout.WriteLine(JsonSerializer.Serialize(new CliResult
        {
            Ok = true,
            OutputPath = output,
            SizeBytes = result.Xlsb.Length,
            SheetCount = result.SheetCount,
            Adjustments = result.Adjustments,
        }, Json));
        return Ok;
    }

    private static int Fail(TextWriter stdout, TextWriter stderr, int code, string message)
    {
        stdout.WriteLine(JsonSerializer.Serialize(new CliResult { Ok = false, Error = message }, Json));
        stderr.WriteLine($"xlsx2xlsb: {message}");
        return code;
    }

    private static void TryDelete(string path)
    {
        try { File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { }
    }

    internal static string Version() =>
        "xlsx2xlsb " + (typeof(Program).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "unknown");
}

internal sealed record CliResult
{
    [JsonPropertyName("ok")] public required bool Ok { get; init; }
    [JsonPropertyName("error")] public string? Error { get; init; }
    [JsonPropertyName("outputPath")] public string? OutputPath { get; init; }
    [JsonPropertyName("sizeBytes")] public int? SizeBytes { get; init; }
    [JsonPropertyName("sheetCount")] public int? SheetCount { get; init; }
    [JsonPropertyName("adjustments")] public ConversionAdjustments? Adjustments { get; init; }
}
