using System.Text.Json;
using XlsxToXlsb.Cli;

namespace XlsxToXlsbLambda.Tests;

public sealed class CliTests : IDisposable
{
    private readonly string _dir = Directory.CreateTempSubdirectory("xlsx2xlsb-tests-").FullName;

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    private (int Code, JsonElement Json, string Stderr) Run(params string[] args)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        int code = Program.Run(args, stdout, stderr);
        var lines = stdout.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Single(lines); // exactly one JSON line, success or failure
        return (code, JsonDocument.Parse(lines[0]).RootElement.Clone(), stderr.ToString());
    }

    private string Input(byte[] content, string name = "in.xlsx")
    {
        var path = Path.Combine(_dir, name);
        File.WriteAllBytes(path, content);
        return path;
    }

    [Fact]
    public void Converts_and_reports_the_result_as_one_json_line()
    {
        var output = Path.Combine(_dir, "nested", "out.xlsb");

        var (code, json, stderr) = Run(Input(SampleWorkbook.Build()), output);

        Assert.Equal(Program.Ok, code);
        Assert.Empty(stderr);
        Assert.True(json.GetProperty("ok").GetBoolean());
        Assert.Equal(output, json.GetProperty("outputPath").GetString());
        Assert.Equal(new FileInfo(output).Length, json.GetProperty("sizeBytes").GetInt32());
        Assert.Equal(2, json.GetProperty("sheetCount").GetInt32());
        Assert.Equal(1, json.GetProperty("adjustments").GetProperty("conditionalFormatsRemoved").GetInt32());
        Assert.False(json.TryGetProperty("error", out _));
        Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(output)!, "*.tmp"));
    }

    [Fact]
    public void Overwrites_an_existing_output()
    {
        var output = Path.Combine(_dir, "out.xlsb");
        File.WriteAllText(output, "stale");

        var (code, _, _) = Run(Input(SampleWorkbook.Build()), output);

        Assert.Equal(Program.Ok, code);
        Assert.NotEqual("stale", File.ReadAllText(output));
    }

    [Fact]
    public void A_non_workbook_fails_with_code_1_and_leaves_no_output()
    {
        var output = Path.Combine(_dir, "out.xlsb");

        var (code, json, stderr) = Run(Input("not a workbook"u8.ToArray()), output);

        Assert.Equal(Program.ConversionFailed, code);
        Assert.False(json.GetProperty("ok").GetBoolean());
        Assert.Contains("not an .xlsx", json.GetProperty("error").GetString());
        Assert.StartsWith("xlsx2xlsb: ", stderr);
        Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(_dir, "*.tmp"));
    }

    [Fact]
    public void A_missing_input_fails_with_code_3()
    {
        var (code, json, _) = Run(Path.Combine(_dir, "missing.xlsx"), Path.Combine(_dir, "out.xlsb"));

        Assert.Equal(Program.IoError, code);
        Assert.Contains("cannot read", json.GetProperty("error").GetString());
    }

    [Theory]
    [InlineData()]
    [InlineData("only-one.xlsx")]
    [InlineData("a.xlsx", "b.xlsb", "c")]
    [InlineData("a.xlsx", "b.xlsx")]
    public void Bad_arguments_fail_with_code_2(params string[] args)
    {
        var (code, json, _) = Run(args);

        Assert.Equal(Program.UsageError, code);
        Assert.False(json.GetProperty("ok").GetBoolean());
    }

    [Fact]
    public void The_output_may_not_be_the_input()
    {
        var path = Input(SampleWorkbook.Build(), "same.xlsb");

        var (code, json, _) = Run(path, path);

        Assert.Equal(Program.UsageError, code);
        Assert.Contains("must differ", json.GetProperty("error").GetString());
    }

    [Fact]
    public void Version_prints_the_tool_name_and_version()
    {
        var stdout = new StringWriter();

        Assert.Equal(Program.Ok, Program.Run(["--version"], stdout, new StringWriter()));
        Assert.StartsWith("xlsx2xlsb 1.0.0", stdout.ToString());
    }
}
