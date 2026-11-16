using Amazon.Lambda.TestUtilities;

namespace XlsxToXlsbLambda.Tests;

public class FunctionTests
{
    private static Function NewFunction() =>
        new(() => throw new InvalidOperationException("S3 must not be touched by an inline request"));

    [Fact]
    public async Task Inline_request_returns_the_xlsb_inline_with_an_xlsb_file_name()
    {
        var request = new ConvertRequest
        {
            FileName = "Evolve-OLA-Sample-Export.xlsx",
            XlsxBase64 = Convert.ToBase64String(SampleWorkbook.Build()),
        };

        var response = await NewFunction().FunctionHandler(request, new TestLambdaContext());

        Assert.Equal("Evolve-OLA-Sample-Export.xlsb", response.FileName);
        Assert.Equal(2, response.SheetCount);
        Assert.NotNull(response.XlsbBase64);
        Assert.Equal(response.SizeBytes, Convert.FromBase64String(response.XlsbBase64!).Length);
        Assert.Null(response.DestinationBucket);
        Assert.Equal(1, response.Adjustments.ConditionalFormatsRemoved);
    }

    [Fact]
    public async Task Inline_request_without_a_file_name_gets_a_default_name()
    {
        var response = await NewFunction().FunctionHandler(
            new ConvertRequest { XlsxBase64 = Convert.ToBase64String(SampleWorkbook.Build()) }, new TestLambdaContext());

        Assert.Equal("workbook.xlsb", response.FileName);
    }

    [Fact]
    public async Task A_request_with_no_input_is_refused()
    {
        var error = await Assert.ThrowsAsync<ArgumentException>(
            () => NewFunction().FunctionHandler(new ConvertRequest(), new TestLambdaContext()));
        Assert.Contains("exactly one input", error.Message);
    }

    [Fact]
    public async Task A_request_with_both_inputs_is_refused()
    {
        var request = new ConvertRequest { XlsxBase64 = "UEs=", SourceBucket = "b", SourceKey = "k.xlsx" };

        await Assert.ThrowsAsync<ArgumentException>(() => NewFunction().FunctionHandler(request, new TestLambdaContext()));
    }

    [Fact]
    public async Task An_s3_request_missing_its_key_is_refused()
    {
        await Assert.ThrowsAsync<ArgumentException>(
            () => NewFunction().FunctionHandler(new ConvertRequest { SourceBucket = "b" }, new TestLambdaContext()));
    }

    [Fact]
    public async Task Invalid_base64_is_refused()
    {
        var error = await Assert.ThrowsAsync<ArgumentException>(
            () => NewFunction().FunctionHandler(new ConvertRequest { XlsxBase64 = "***" }, new TestLambdaContext()));
        Assert.Contains("base64", error.Message);
    }

    [Fact]
    public async Task A_non_workbook_is_refused_with_a_conversion_error()
    {
        var request = new ConvertRequest { XlsxBase64 = Convert.ToBase64String("hello"u8.ToArray()) };

        await Assert.ThrowsAsync<XlsbConversionException>(() => NewFunction().FunctionHandler(request, new TestLambdaContext()));
    }

    [Fact]
    public void The_runtime_constructor_does_not_create_an_s3_client_eagerly()
    {
        // Constructing must not need credentials or a region; the client is created on first S3 use.
        var exception = Record.Exception(() => new Function());
        Assert.Null(exception);
    }
}
