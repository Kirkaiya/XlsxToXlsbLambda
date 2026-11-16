using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxToXlsbLambda.Tests;

/// <summary>
/// Builds a small synthetic .xlsx carrying every construct the OLA export has that OfficeIMO's XLSB
/// writer rejects: conditional formatting with dxf styles, an indexed-color palette, overlapping column
/// ranges and a sort state inside the AutoFilter. No customer data.
/// </summary>
internal static class SampleWorkbook
{
    public static readonly string[] Headers = ["Server Name", "vCPU", "RAM GB", "Instance Type"];

    public static readonly object[][] Rows =
    [
        ["app-01", 4, 16.5, "m7i.xlarge"],
        ["db-02", 8, 64, "r7i.2xlarge"],
        ["web-03", 2, 8.25, "t3.large"],
    ];

    public static byte[] Build(bool withQuirks = true)
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets());

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = BuildStylesheet(withQuirks);

            AddSheet(workbookPart, 1, "EC2 Shared (Rightsized)", withQuirks);
            AddSheet(workbookPart, 2, "Blueprint", withQuirks: false);
            workbookPart.Workbook.Save();
        }
        return stream.ToArray();
    }

    private static void AddSheet(WorkbookPart workbookPart, uint sheetId, string name, bool withQuirks)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var data = new SheetData();
        data.Append(Row(1, Headers.Cast<object>().ToArray()));
        uint rowIndex = 2;
        foreach (var values in Rows)
            data.Append(Row(rowIndex++, values));

        var worksheet = new Worksheet();
        if (withQuirks)
        {
            // A range, then one column inside it: what the toolkit writes.
            worksheet.Append(new Columns(
                new Column { Min = 1, Max = 4, Width = 12, CustomWidth = true },
                new Column { Min = 2, Max = 2, Width = 30, CustomWidth = true }));
        }
        worksheet.Append(data);
        string lastRow = $"D{Rows.Length + 1}";
        if (withQuirks)
        {
            worksheet.Append(new AutoFilter(
                new SortState(new SortCondition { Reference = "A2" }) { Reference = $"A2:{lastRow}" })
            { Reference = $"A1:{lastRow}" });
            worksheet.Append(new ConditionalFormatting(
                new ConditionalFormattingRule(new Formula("100")) { Type = ConditionalFormatValues.CellIs, Operator = ConditionalFormattingOperatorValues.GreaterThan, FormatId = 0, Priority = 1 })
            { SequenceOfReferences = new ListValue<StringValue> { InnerText = $"B2:B{Rows.Length + 1}" } });
        }
        worksheetPart.Worksheet = worksheet;

        workbookPart.Workbook!.Sheets!.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = name,
        });
    }

    private static Row Row(uint index, object[] values)
    {
        var row = new Row { RowIndex = index };
        for (int i = 0; i < values.Length; i++)
        {
            string reference = $"{(char)('A' + i)}{index}";
            row.Append(values[i] switch
            {
                string text => new Cell { CellReference = reference, DataType = CellValues.InlineString, InlineString = new InlineString(new Text(text)) },
                IConvertible number => new Cell { CellReference = reference, CellValue = new CellValue(System.Convert.ToDouble(number)) },
                _ => throw new ArgumentException("unsupported sample value"),
            });
        }
        return row;
    }

    private static Stylesheet BuildStylesheet(bool withQuirks)
    {
        var stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(new CellFormat()) { Count = 1 });
        if (withQuirks)
        {
            stylesheet.Append(new DifferentialFormats(
                new DifferentialFormat(new Fill(new PatternFill(new BackgroundColor { Rgb = "FFFFC7CE" }))))
            { Count = 1 });
            stylesheet.Append(new Colors(new IndexedColors(
                new RgbColor { Rgb = "00000000" }, new RgbColor { Rgb = "00FFFFFF" })));
        }
        return stylesheet;
    }
}
