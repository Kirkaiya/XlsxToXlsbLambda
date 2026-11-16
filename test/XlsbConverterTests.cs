using XlsxToXlsb;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;

namespace XlsxToXlsbLambda.Tests;

public class XlsbConverterTests
{
    [Fact]
    public void Converts_a_workbook_with_every_known_quirk_and_keeps_the_values()
    {
        var result = XlsbConverter.Convert(SampleWorkbook.Build(withQuirks: true));

        Assert.Equal(2, result.SheetCount);
        AssertIsXlsb(result.Xlsb);
        AssertSampleValues(result.Xlsb);
        Assert.Equal(new ConversionAdjustments
        {
            ConditionalFormatsRemoved = 1,
            DifferentialFormatsRemoved = 1,
            IndexedColorPaletteRemoved = true,
            StyleExtensionsRemoved = false,
            WorksheetsWithColumnRangesNormalized = 1,
            AutoFilterSortStatesRemoved = 1,
        }, result.Adjustments);
    }

    [Fact]
    public void A_workbook_without_quirks_needs_no_adjustments()
    {
        var result = XlsbConverter.Convert(SampleWorkbook.Build(withQuirks: false));

        AssertSampleValues(result.Xlsb);
        Assert.Equal(new ConversionAdjustments(), result.Adjustments);
    }

    [Fact]
    public void Overlapping_column_ranges_become_disjoint_and_the_later_definition_wins()
    {
        var columns = new Columns(
            new Column { Min = 1, Max = 5, Width = 10 },
            new Column { Min = 3, Max = 3, Width = 30 },
            new Column { Min = 5, Max = 16384, Width = 8 });

        Assert.True(XlsbConverter.NormalizeColumns(columns));

        var ranges = columns.Elements<Column>().Select(c => (c.Min!.Value, c.Max!.Value, c.Width!.Value)).ToList();
        Assert.Equal([(1u, 2u, 10d), (3u, 3u, 30d), (4u, 4u, 10d), (5u, 16384u, 8d)], ranges);
    }

    [Fact]
    public void Ordered_disjoint_column_ranges_are_left_alone()
    {
        var columns = new Columns(new Column { Min = 1, Max = 2, Width = 10 }, new Column { Min = 3, Max = 9, Width = 12 });

        Assert.False(XlsbConverter.NormalizeColumns(columns));
        Assert.Equal(2, columns.Elements<Column>().Count());
    }

    [Fact]
    public void Input_that_is_not_a_zip_package_is_refused()
    {
        var error = Assert.Throws<XlsbConversionException>(() => XlsbConverter.Convert("not a workbook"u8.ToArray()));
        Assert.Contains("not an .xlsx", error.Message);
    }

    [Fact]
    public void A_zip_that_is_not_a_workbook_is_refused()
    {
        using var zip = new MemoryStream();
        using (var archive = new System.IO.Compression.ZipArchive(zip, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            archive.CreateEntry("hello.txt");

        Assert.Throws<XlsbConversionException>(() => XlsbConverter.Convert(zip.ToArray()));
    }

    /// <summary>
    /// Converts a real OLA export when <c>OLA_SAMPLE_XLSX</c> points at one, and checks every cell's
    /// text survives. Skipped otherwise, because the repo carries no customer workbooks.
    /// </summary>
    [RealWorkbookFact]
    public void Converts_a_real_ola_export_cell_for_cell()
    {
        var path = Environment.GetEnvironmentVariable(RealWorkbookFactAttribute.Variable)!;
        var xlsx = File.ReadAllBytes(path);

        var result = XlsbConverter.Convert(xlsx);

        AssertIsXlsb(result.Xlsb);
        using var source = ExcelDocument.Load(new MemoryStream(xlsx));
        using var converted = ExcelDocument.Load(new MemoryStream(result.Xlsb));
        Assert.Equal(source.Sheets.Select(s => s.Name), converted.Sheets.Select(s => s.Name));

        int compared = 0;
        foreach (var (before, after) in source.Sheets.Zip(converted.Sheets))
        {
            // The larger of the two used ranges, so a cell that appears or vanishes is caught too.
            var (rowsBefore, colsBefore) = UsedExtent(before);
            var (rowsAfter, colsAfter) = UsedExtent(after);
            int rows = Math.Max(rowsBefore, rowsAfter), cols = Math.Max(colsBefore, colsAfter);
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                {
                    var expected = StoredValue(before, r, c);
                    var actual = StoredValue(after, r, c);
                    Assert.True(expected == actual, $"{before.Name}!R{r}C{c}: '{expected}' became '{actual}'");
                    if (!string.IsNullOrEmpty(expected)) compared++;
                }
        }
        Assert.True(compared > 0, "the workbook had no cells to compare");
    }

    /// <summary>
    /// The cell's value for comparison: numbers to 15 significant digits, anything else as its text.
    /// OfficeIMO's .xlsb reader reports numbers at 15 significant digits (74.78400000000001 reads back
    /// as 74.784), so a finer comparison through it would fail on its reader, not on the file. The
    /// stored doubles were checked bit-for-bit with an independent reader (python-calamine) — see README.
    /// </summary>
    private static string? StoredValue(ExcelSheet sheet, int row, int column)
    {
        if (!sheet.TryGetCellValueSnapshot(row, column, out var info) || info is null) return null;
        var invariant = System.Globalization.CultureInfo.InvariantCulture;
        if (info.Kind.ToString() == "Number" &&
            double.TryParse(System.Convert.ToString(info.RawValue, invariant), System.Globalization.NumberStyles.Float, invariant, out var number))
            return number.ToString("G15", invariant);
        return string.IsNullOrEmpty(info.Text) ? null : info.Text;
    }

    private static (int Rows, int Cols) UsedExtent(ExcelSheet sheet)
    {
        var used = sheet.UsedRange;
        return used is null ? (0, 0) : (used.LastRow, used.LastColumn);
    }

    private static void AssertIsXlsb(byte[] bytes)
    {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        Assert.Contains(archive.Entries, e => e.FullName == "xl/workbook.bin");
        Assert.DoesNotContain(archive.Entries, e => e.FullName == "xl/workbook.xml");
    }

    private static void AssertSampleValues(byte[] xlsb)
    {
        using var document = ExcelDocument.Load(new MemoryStream(xlsb));
        Assert.Equal(["EC2 Shared (Rightsized)", "Blueprint"], document.Sheets.Select(s => s.Name));
        foreach (var sheet in document.Sheets)
        {
            for (int c = 0; c < SampleWorkbook.Headers.Length; c++)
            {
                Assert.True(sheet.TryGetCellText(1, c + 1, out var header));
                Assert.Equal(SampleWorkbook.Headers[c], header);
            }
            for (int r = 0; r < SampleWorkbook.Rows.Length; r++)
                for (int c = 0; c < SampleWorkbook.Rows[r].Length; c++)
                {
                    Assert.True(sheet.TryGetCellText(r + 2, c + 1, out var text), $"{sheet.Name}!R{r + 2}C{c + 1} missing");
                    Assert.Equal(System.Convert.ToString(SampleWorkbook.Rows[r][c], System.Globalization.CultureInfo.InvariantCulture), text);
                }
        }
    }
}

/// <summary>A fact that runs only when <see cref="Variable"/> names an existing .xlsx.</summary>
public sealed class RealWorkbookFactAttribute : FactAttribute
{
    public const string Variable = "OLA_SAMPLE_XLSX";

    public RealWorkbookFactAttribute()
    {
        var path = Environment.GetEnvironmentVariable(Variable);
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            Skip = $"Set {Variable} to a real OLA export .xlsx to run this test.";
    }
}
