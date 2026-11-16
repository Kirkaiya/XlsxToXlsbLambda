using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO;
using OfficeIMO.Excel;

namespace XlsxToXlsbLambda;

/// <summary>What the converter had to change before OfficeIMO would write the workbook as .xlsb.</summary>
public sealed record ConversionAdjustments
{
    /// <summary>Conditional-formatting blocks removed (cosmetic highlighting; cell values are untouched).</summary>
    [JsonPropertyName("conditionalFormatsRemoved")]
    public int ConditionalFormatsRemoved { get; init; }

    /// <summary>Differential-format (dxf) style records removed — they exist only to serve conditional formatting.</summary>
    [JsonPropertyName("differentialFormatsRemoved")]
    public int DifferentialFormatsRemoved { get; init; }

    /// <summary>True when a custom indexed-color palette was removed (Excel falls back to its default palette).</summary>
    [JsonPropertyName("indexedColorPaletteRemoved")]
    public bool IndexedColorPaletteRemoved { get; init; }

    /// <summary>True when a stylesheet extension list was removed.</summary>
    [JsonPropertyName("styleExtensionsRemoved")]
    public bool StyleExtensionsRemoved { get; init; }

    /// <summary>Worksheets whose overlapping or unordered column ranges were rewritten (the later definition of a column wins).</summary>
    [JsonPropertyName("worksheetsWithColumnRangesNormalized")]
    public int WorksheetsWithColumnRangesNormalized { get; init; }

    /// <summary>Saved sort states removed from worksheet AutoFilters (the rows keep their current order; the filter stays).</summary>
    [JsonPropertyName("autoFilterSortStatesRemoved")]
    public int AutoFilterSortStatesRemoved { get; init; }
}

/// <summary>The converted workbook and what was changed to produce it.</summary>
public sealed record ConversionResult(byte[] Xlsb, int SheetCount, ConversionAdjustments Adjustments);

/// <summary>
/// Converts an .xlsx workbook to the Excel binary format (.xlsb) with OfficeIMO.Excel.
///
/// OfficeIMO's native XLSB writer (3.6.x) refuses several constructs it cannot encode yet, and the
/// ecs-etl-toolkit OLA export contains all of them. Before saving, the workbook is normalized in
/// memory: conditional formatting and its dxf styles, the indexed-color palette and stylesheet
/// extensions are removed; overlapping column ranges are rewritten; AutoFilter sort states are
/// dropped. None of these change a cell value. Anything else OfficeIMO cannot encode surfaces as
/// <see cref="XlsbConversionException"/> rather than being silently dropped.
/// </summary>
public static class XlsbConverter
{
    public static ConversionResult Convert(byte[] xlsx)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        if (xlsx.Length < 4 || xlsx[0] != 'P' || xlsx[1] != 'K')
            throw new XlsbConversionException("The input is not an .xlsx workbook (it is not a ZIP package).");

        using var working = new MemoryStream();
        working.Write(xlsx);
        working.Position = 0;

        ConversionAdjustments adjustments;
        try
        {
            adjustments = Normalize(working);
        }
        catch (Exception e) when (e is OpenXmlPackageException or InvalidDataException or FileFormatException)
        {
            throw new XlsbConversionException($"The input could not be opened as an .xlsx workbook: {e.Message}", e);
        }

        working.Position = 0;
        try
        {
            using var document = ExcelDocument.Load(working, null, out OfficeConversionReport loadReport);
            if (loadReport.HasLoss)
                throw new XlsbConversionException(
                    "OfficeIMO reported content loss while reading the workbook: " +
                    string.Join("; ", loadReport.Diagnostics));

            using var output = new MemoryStream();
            document.Save(output, new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
            return new ConversionResult(output.ToArray(), document.Sheets.Count, adjustments);
        }
        catch (NotSupportedException e)
        {
            throw new XlsbConversionException($"OfficeIMO cannot write this workbook as .xlsb: {e.Message}", e);
        }
    }

    /// <summary>Rewrites the package in place so OfficeIMO's XLSB writer accepts it.</summary>
    private static ConversionAdjustments Normalize(Stream package)
    {
        using var document = SpreadsheetDocument.Open(package, true);
        var workbookPart = document.WorkbookPart
            ?? throw new XlsbConversionException("The input has no workbook part.");

        int conditionalFormats = 0, columnSheets = 0, sortStates = 0;
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var worksheet = worksheetPart.Worksheet;
            if (worksheet is null) continue;

            foreach (var block in worksheet.Elements<ConditionalFormatting>().ToList())
            {
                block.Remove();
                conditionalFormats++;
            }

            foreach (var autoFilter in worksheet.Elements<AutoFilter>())
                foreach (var sortState in autoFilter.Elements<SortState>().ToList())
                {
                    sortState.Remove();
                    sortStates++;
                }

            bool normalized = false;
            foreach (var columns in worksheet.Elements<Columns>())
                normalized |= NormalizeColumns(columns);
            if (normalized) columnSheets++;
        }

        int dxfs = 0;
        bool palette = false, extensions = false;
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        if (stylesheet is not null)
        {
            if (stylesheet.DifferentialFormats is { } differentialFormats)
            {
                dxfs = differentialFormats.Elements<DifferentialFormat>().Count();
                differentialFormats.Remove();
            }
            if (stylesheet.Colors is { } colors)
            {
                colors.Remove();
                palette = true;
            }
            if (stylesheet.StylesheetExtensionList is { } extensionList)
            {
                extensionList.Remove();
                extensions = true;
            }
        }

        return new ConversionAdjustments
        {
            ConditionalFormatsRemoved = conditionalFormats,
            DifferentialFormatsRemoved = dxfs,
            IndexedColorPaletteRemoved = palette,
            StyleExtensionsRemoved = extensions,
            WorksheetsWithColumnRangesNormalized = columnSheets,
            AutoFilterSortStatesRemoved = sortStates,
        };
    }

    /// <summary>
    /// Makes column ranges ordered and non-overlapping. A writer that sets a range and then one column
    /// inside it (the toolkit does) means the later setting for that column, so the later one wins.
    /// Runs of adjacent columns that came from the same original definition are kept as one range.
    /// </summary>
    /// <returns>True when the ranges were rewritten.</returns>
    internal static bool NormalizeColumns(Columns columns)
    {
        var original = columns.Elements<Column>().ToList();
        if (IsOrderedAndDisjoint(original)) return false;

        var owner = new SortedDictionary<uint, Column>();
        foreach (var column in original)
        {
            if (column.Min?.Value is not uint min || column.Max?.Value is not uint max || min > max) continue;
            for (uint index = min; index <= max; index++)
                owner[index] = column;
        }

        var rewritten = new List<Column>();
        Column? current = null, currentSource = null;
        foreach (var (index, source) in owner)
        {
            if (current is not null && ReferenceEquals(source, currentSource) && current.Max!.Value == index - 1)
            {
                current.Max = index;
                continue;
            }
            current = (Column)source.CloneNode(true);
            current.Min = index;
            current.Max = index;
            currentSource = source;
            rewritten.Add(current);
        }

        columns.RemoveAllChildren<Column>();
        foreach (var column in rewritten)
            columns.Append(column);
        return true;
    }

    private static bool IsOrderedAndDisjoint(IReadOnlyList<Column> columns)
    {
        uint previousMax = 0;
        foreach (var column in columns)
        {
            if (column.Min?.Value is not uint min || column.Max?.Value is not uint max || min > max) return false;
            if (min <= previousMax) return false;
            previousMax = max;
        }
        return true;
    }
}

/// <summary>The input cannot be converted; the message says why.</summary>
public sealed class XlsbConversionException : Exception
{
    public XlsbConversionException(string message) : base(message) { }
    public XlsbConversionException(string message, Exception inner) : base(message, inner) { }
}
