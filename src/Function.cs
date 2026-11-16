using System.Text.Json.Serialization;
using Amazon.Lambda.Core;
using Amazon.S3;
using Amazon.S3.Model;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace XlsxToXlsbLambda;

/// <summary>
/// The conversion request. Supply the workbook one of two ways:
///   inline:  <c>xlsxBase64</c> (keep the file under ~4 MB; Lambda's synchronous payload limit is 6 MB
///            and base64 adds a third), or
///   S3:      <c>sourceBucket</c> + <c>sourceKey</c>.
/// Set <c>destinationBucket</c> to have the .xlsb written to S3 instead of returned inline;
/// <c>destinationKey</c> defaults to the source key (or file name) with an .xlsb extension.
/// </summary>
public sealed record ConvertRequest
{
    [JsonPropertyName("fileName")] public string? FileName { get; init; }
    [JsonPropertyName("xlsxBase64")] public string? XlsxBase64 { get; init; }
    [JsonPropertyName("sourceBucket")] public string? SourceBucket { get; init; }
    [JsonPropertyName("sourceKey")] public string? SourceKey { get; init; }
    [JsonPropertyName("destinationBucket")] public string? DestinationBucket { get; init; }
    [JsonPropertyName("destinationKey")] public string? DestinationKey { get; init; }
}

public sealed record ConvertResponse
{
    [JsonPropertyName("fileName")] public required string FileName { get; init; }
    [JsonPropertyName("sizeBytes")] public required int SizeBytes { get; init; }
    [JsonPropertyName("sheetCount")] public required int SheetCount { get; init; }

    /// <summary>The .xlsb, when no destination bucket was given.</summary>
    [JsonPropertyName("xlsbBase64")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? XlsbBase64 { get; init; }

    [JsonPropertyName("destinationBucket")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DestinationBucket { get; init; }

    [JsonPropertyName("destinationKey")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DestinationKey { get; init; }

    [JsonPropertyName("adjustments")] public required ConversionAdjustments Adjustments { get; init; }
}

public class Function
{
    /// <summary>Base64 output above this is refused inline: the synchronous response limit is 6 MB, JSON included.</summary>
    internal const int MaxInlineBase64Chars = 5_900_000;

    private const string XlsbContentType = "application/vnd.ms-excel.sheet.binary.macroEnabled.12";

    private readonly Lazy<IAmazonS3> _s3;

    /// <summary>Used by the Lambda runtime.</summary>
    public Function() : this(() => new AmazonS3Client()) { }

    /// <summary>Used by tests to supply an S3 client.</summary>
    public Function(Func<IAmazonS3> s3Factory) => _s3 = new Lazy<IAmazonS3>(s3Factory);

    /// <summary>Converts one .xlsx workbook to .xlsb. Failures throw, so the caller sees a Lambda function error with the reason.</summary>
    public async Task<ConvertResponse> FunctionHandler(ConvertRequest request, ILambdaContext context)
    {
        ArgumentNullException.ThrowIfNull(request);
        bool inline = !string.IsNullOrWhiteSpace(request.XlsxBase64);
        bool fromS3 = !string.IsNullOrWhiteSpace(request.SourceBucket) || !string.IsNullOrWhiteSpace(request.SourceKey);
        if (inline == fromS3)
            throw new ArgumentException("Supply exactly one input: 'xlsxBase64', or 'sourceBucket' with 'sourceKey'.");

        byte[] xlsx;
        string sourceName;
        if (inline)
        {
            try
            {
                xlsx = Convert.FromBase64String(request.XlsxBase64!);
            }
            catch (FormatException)
            {
                throw new ArgumentException("'xlsxBase64' is not valid base64.");
            }
            sourceName = string.IsNullOrWhiteSpace(request.FileName) ? "workbook.xlsx" : request.FileName!;
        }
        else
        {
            if (string.IsNullOrWhiteSpace(request.SourceBucket) || string.IsNullOrWhiteSpace(request.SourceKey))
                throw new ArgumentException("S3 input needs both 'sourceBucket' and 'sourceKey'.");
            xlsx = await ReadObjectAsync(request.SourceBucket!, request.SourceKey!);
            sourceName = string.IsNullOrWhiteSpace(request.FileName) ? request.SourceKey! : request.FileName!;
        }

        context.Logger.LogInformation($"Converting '{sourceName}' ({xlsx.Length:N0} bytes)");
        var started = DateTime.UtcNow;
        var result = XlsbConverter.Convert(xlsx);
        var outputName = Path.ChangeExtension(Path.GetFileName(sourceName), ".xlsb");
        context.Logger.LogInformation(
            $"Converted '{sourceName}' -> '{outputName}' ({result.Xlsb.Length:N0} bytes, {result.SheetCount} sheets) " +
            $"in {(DateTime.UtcNow - started).TotalMilliseconds:N0} ms; adjustments: {result.Adjustments}");

        if (!string.IsNullOrWhiteSpace(request.DestinationBucket))
        {
            var key = !string.IsNullOrWhiteSpace(request.DestinationKey)
                ? request.DestinationKey!
                : Path.ChangeExtension(inline ? outputName : request.SourceKey!, ".xlsb");
            await _s3.Value.PutObjectAsync(new PutObjectRequest
            {
                BucketName = request.DestinationBucket,
                Key = key,
                InputStream = new MemoryStream(result.Xlsb),
                ContentType = XlsbContentType,
            });
            return new ConvertResponse
            {
                FileName = outputName,
                SizeBytes = result.Xlsb.Length,
                SheetCount = result.SheetCount,
                DestinationBucket = request.DestinationBucket,
                DestinationKey = key,
                Adjustments = result.Adjustments,
            };
        }

        var base64 = Convert.ToBase64String(result.Xlsb);
        if (base64.Length > MaxInlineBase64Chars)
            throw new InvalidOperationException(
                $"The .xlsb is {result.Xlsb.Length:N0} bytes, too large to return inline (Lambda's 6 MB response limit). " +
                "Set 'destinationBucket' to have it written to S3.");

        return new ConvertResponse
        {
            FileName = outputName,
            SizeBytes = result.Xlsb.Length,
            SheetCount = result.SheetCount,
            XlsbBase64 = base64,
            Adjustments = result.Adjustments,
        };
    }

    private async Task<byte[]> ReadObjectAsync(string bucket, string key)
    {
        using var response = await _s3.Value.GetObjectAsync(bucket, key);
        using var buffer = new MemoryStream();
        await response.ResponseStream.CopyToAsync(buffer);
        return buffer.ToArray();
    }
}
