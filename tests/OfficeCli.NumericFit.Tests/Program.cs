// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using OfficeCli.Handlers;

const string NumericOverflowSubtype = "numeric_overflow";

var standardPath = Path.Combine(Path.GetTempPath(), $"officecli-numeric-fit-{Guid.NewGuid():N}.xlsx");
var date1904Path = Path.Combine(Path.GetTempPath(), $"officecli-numeric-fit-1904-{Guid.NewGuid():N}.xlsx");
try
{
    CreateStandardFixture(standardPath);
    VerifyStandardWorkbook(standardPath);

    CreateDate1904Fixture(date1904Path);
    VerifyDate1904Workbook(date1904Path);

    Console.WriteLine("XLSX numeric-fit issue tests passed.");
}
finally
{
    if (File.Exists(standardPath)) File.Delete(standardPath);
    if (File.Exists(date1904Path)) File.Delete(date1904Path);
}

static void VerifyStandardWorkbook(string path)
{
    string[] expectedPaths =
    [
        "/FormulaView/A1", // Show Formulas does not hide an ordinary number
        "/Sheet1/A1",     // explicit narrow number
        "/Sheet1/B1",     // explicit narrow date serial
        "/Sheet1/F1",     // wrapText does not make a number spill
        "/Sheet1/M1",     // evaluated numeric formula
        "/Sheet1/N1",     // ISO t=d date
        "/Sheet1/P1",     // bracketed color numeric format
        "/Sheet1/T1",     // inherited column style
        "/Sheet1/U3",     // row style wins over column style
        "/Sheet1/W1",     // direct cell style wins over column shrinkToFit
        "/Sheet1/X4",     // direct cell style wins over row shrinkToFit
        "/Sheet1/AD1",    // active non-General section is still checked
    ];

    using var handler = new ExcelHandler(path, editable: false);
    var allIssues = handler.ViewAsIssues();
    var numericIssues = allIssues
        .Where(issue => issue.Subtype == NumericOverflowSubtype)
        .ToList();

    AssertPaths(expectedPaths, numericIssues.Select(issue => issue.Path), "default issues scan");

    foreach (var issue in numericIssues)
    {
        Assert(issue.Type == IssueType.Format, $"{issue.Path} should use the Format bucket");
        Assert(issue.Severity == IssueSeverity.Warning, $"{issue.Path} should be a warning");
        Assert(issue.Message.Contains("numeric overflow", StringComparison.Ordinal),
            $"{issue.Path} should describe the rendering defect");
        Assert(issue.Suggestion?.Contains("suggest.width=", StringComparison.Ordinal) == true,
            $"{issue.Path} should carry an actionable width suggestion");
    }
    Assert(numericIssues.Single(issue => issue.Path == "/Sheet1/A1").Message
            .Contains("'1,234,567.89'", StringComparison.Ordinal),
        "numeric finding should use the formatted display value");
    Assert(numericIssues.Single(issue => issue.Path == "/Sheet1/B1").Message
            .Contains("'2024-10-02'", StringComparison.Ordinal),
        "date finding should use the formatted display value");
    Assert(numericIssues.Single(issue => issue.Path == "/Sheet1/N1").Message
            .Contains("'2024-10-02'", StringComparison.Ordinal),
        "ISO date finding should use its absolute date");
    Assert(numericIssues.All(issue => issue.Path is not "/Sheet1/AE1"
            and not "/Sheet1/AF1" and not "/Sheet1/AG1"),
        "non-finite numeric values should not produce width findings");
    Assert(numericIssues.All(issue => issue.Path != "/Sheet1/AH1"),
        "negative elapsed values in a 1900 workbook need a non-width fix");

    var exactIssues = handler.ViewAsIssues(NumericOverflowSubtype);
    AssertPaths(expectedPaths, exactIssues.Select(issue => issue.Path), "exact subtype filter");

    var formatIssues = handler.ViewAsIssues("format")
        .Where(issue => issue.Subtype == NumericOverflowSubtype);
    AssertPaths(expectedPaths, formatIssues.Select(issue => issue.Path), "format bucket filter");

    // The scanner receives only the remaining capacity and stops at the first
    // matching Format finding; the broken defined name cannot consume it.
    var exactLimited = handler.ViewAsIssues(NumericOverflowSubtype, limit: 1);
    AssertPaths(["/Sheet1/A1"], exactLimited.Select(issue => issue.Path), "exact subtype limit");
    var formatLimited = handler.ViewAsIssues("format", limit: 1);
    AssertPaths(["/Sheet1/A1"], formatLimited.Select(issue => issue.Path), "format bucket limit");
}

static void VerifyDate1904Workbook(string path)
{
    using var handler = new ExcelHandler(path, editable: false);
    var issues = handler.ViewAsIssues(NumericOverflowSubtype);
    AssertPaths(
        [
            "/Sheet1/A1", "/Sheet1/C1", "/Sheet1/E1", "/Sheet1/F1",
            "/Sheet1/G1", "/Sheet1/H1", "/Sheet1/I1"
        ],
        issues.Select(issue => issue.Path),
        "1904 date-system scan");
    Assert(issues.Single(issue => issue.Path == "/Sheet1/A1").Message
            .Contains("'1904-01-01'", StringComparison.Ordinal),
        "a stored 1904 serial should use the workbook epoch");
    Assert(issues.All(issue => issue.Path is not "/Sheet1/B1" and not "/Sheet1/D1"),
        "date-formatted formulas in a 1904 workbook should be skipped conservatively");
    Assert(issues.Single(issue => issue.Path == "/Sheet1/F1").Message
            .Contains("'36:00:00'", StringComparison.Ordinal),
        "1904 workbooks must not epoch-shift positive elapsed hours");
    Assert(issues.Single(issue => issue.Path == "/Sheet1/G1").Message
            .Contains("'-36:00:00'", StringComparison.Ordinal),
        "1904 workbooks must not epoch-shift negative elapsed hours");
}

static void CreateStandardFixture(string path)
{
    using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = CreateStylesheet();

    var visiblePart = workbookPart.AddNewPart<WorksheetPart>();
    visiblePart.Worksheet = CreateVisibleWorksheet();
    var formulaViewPart = workbookPart.AddNewPart<WorksheetPart>();
    formulaViewPart.Worksheet = CreateFormulaViewWorksheet();
    var hiddenPart = workbookPart.AddNewPart<WorksheetPart>();
    hiddenPart.Worksheet = CreateHiddenWorksheet();

    workbookPart.Workbook.Append(
        new Sheets(
            SheetFor(workbookPart, visiblePart, 1, "Sheet1"),
            SheetFor(workbookPart, formulaViewPart, 2, "FormulaView"),
            SheetFor(workbookPart, hiddenPart, 3, "HiddenSheet", SheetStateValues.Hidden)),
        new DefinedNames(new DefinedName("#REF!") { Name = "BrokenName" }));
    workbookPart.Workbook.Save();
}

static void CreateDate1904Fixture(string path)
{
    using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(new WorkbookProperties { Date1904 = true });
    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = CreateStylesheet();
    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart.Worksheet = new Worksheet(
        new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 8.43D },
        new Columns(
            Col(1, 4), Col(2, 4), Col(3, 2), Col(4, 4), Col(5, 2),
            Col(6, 2), Col(7, 2), Col(8, 2), Col(9, 2)),
        new SheetData(new Row(
            NumberCell("A1", "0", 2),
            FormulaCell("B1", "DATE(2024,10,2)", 2),
            FormulaCell("C1", "1234567.89", 1),
            FormulaCell("D1", "A1", 2),
            NumberCell("E1", "0.5", 10),
            NumberCell("F1", "1.5", 11),
            NumberCell("G1", "-1.5", 11),
            NumberCell("H1", "0.000011574074", 12),
            NumberCell("I1", "1234567", 13))
        { RowIndex = 1 }));
    workbookPart.Workbook.Append(new Sheets(
        SheetFor(workbookPart, worksheetPart, 1, "Sheet1")));
    workbookPart.Workbook.Save();
}

static Sheet SheetFor(
    WorkbookPart workbookPart,
    WorksheetPart worksheetPart,
    uint id,
    string name,
    SheetStateValues? state = null)
{
    var sheet = new Sheet
    {
        Id = workbookPart.GetIdOfPart(worksheetPart),
        SheetId = id,
        Name = name
    };
    if (state != null) sheet.State = state.Value;
    return sheet;
}

static Stylesheet CreateStylesheet()
{
    var numberingFormats = new NumberingFormats(
        new NumberingFormat { NumberFormatId = 164, FormatCode = "#,##0.00" },
        new NumberingFormat { NumberFormatId = 165, FormatCode = "yyyy-mm-dd" },
        new NumberingFormat { NumberFormatId = 166, FormatCode = "[Red]General" },
        new NumberingFormat { NumberFormatId = 167, FormatCode = "[Red]#,##0.00" },
        new NumberingFormat { NumberFormatId = 168, FormatCode = "General;0.00" })
    { Count = 5 };

    var fonts = new Fonts(
        new Font(new FontSize { Val = 11D }, new FontName { Val = "Calibri" }))
    { Count = 1 };
    var fills = new Fills(
        new Fill(new PatternFill { PatternType = PatternValues.None }),
        new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
    { Count = 2 };
    var borders = new Borders(new Border()) { Count = 1 };
    var cellStyleFormats = new CellStyleFormats(
        new CellFormat(),
        new CellFormat
        {
            ApplyAlignment = true,
            Alignment = new Alignment { ShrinkToFit = true }
        })
    { Count = 2 };
    var cellFormats = new CellFormats(
        new CellFormat(),
        new CellFormat { NumberFormatId = 164, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 165, FontId = 0, ApplyNumberFormat = true },
        new CellFormat
        {
            NumberFormatId = 164,
            FontId = 0,
            ApplyNumberFormat = true,
            ApplyAlignment = true,
            Alignment = new Alignment { ShrinkToFit = true }
        },
        new CellFormat
        {
            NumberFormatId = 164,
            FontId = 0,
            ApplyNumberFormat = true,
            ApplyAlignment = true,
            Alignment = new Alignment { WrapText = true }
        },
        new CellFormat { NumberFormatId = 199, FontId = 0, ApplyNumberFormat = true },
        new CellFormat
        {
            NumberFormatId = 164,
            FontId = 0,
            FormatId = 1,
            ApplyNumberFormat = true
        },
        new CellFormat { NumberFormatId = 166, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 167, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 168, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 45, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 46, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 47, FontId = 0, ApplyNumberFormat = true },
        new CellFormat { NumberFormatId = 48, FontId = 0, ApplyNumberFormat = true })
    { Count = 14 };
    var cellStyles = new CellStyles(
        new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 },
        new CellStyle { Name = "ShrinkBase", FormatId = 1 })
    { Count = 2 };

    return new Stylesheet(
        numberingFormats,
        fonts,
        fills,
        borders,
        cellStyleFormats,
        cellFormats,
        cellStyles,
        new DifferentialFormats { Count = 0 },
        new TableStyles { Count = 0 });
}

static Worksheet CreateVisibleWorksheet()
{
    var columns = new Columns(
        Col(1, 2), Col(2, 4), Col(3, 24), Col(4, 2), Col(5, 2), Col(6, 2),
        Col(7, 8), Col(8, 8), Col(9, 2, hidden: true), Col(10, 2),
        Col(11, 2), Col(12, 2), Col(13, 2), Col(14, 4), Col(15, 2), Col(16, 2),
        Col(17, 7), Col(18, 7),
        Col(20, 2, styleIndex: 1), Col(21, 2, styleIndex: 3),
        Col(22, 2, styleIndex: 1), Col(23, 2, styleIndex: 3), Col(24, 2),
        Col(25, 2), Col(26, 2, styleIndex: 3), Col(27, 2, styleIndex: 1), Col(28, 2),
        Col(29, 2), Col(30, 2), Col(31, 2), Col(32, 2), Col(33, 2), Col(34, 2),
        Col(201, 7), Col(202, 7));

    var row1 = new Row(
        NumberCell("A1", "1234567.89", 1),
        NumberCell("B1", "45567", 2),
        NumberCell("C1", "1234567.89", 1),
        InlineTextCell("D1", "1234567.89"),
        NumberCell("E1", "1234567.89", 3),
        NumberCell("F1", "1234567.89", 4),
        NumberCell("G1", "1234567.89", 1),
        NumberCell("H1", "9876543.21", 1),
        NumberCell("I1", "1234567.89", 1),
        NumberCell("K1", "123456789", 0),
        NumberCell("L1", "1234567.89", 5),
        FormulaCell("M1", "1234567.89", 1),
        IsoDateCell("N1", "2024-10-02T00:00:00Z", 2),
        NumberCell("O1", "123456789", 7),
        NumberCell("P1", "-1234567.89", 8),
        NumberCell("T1", "1234567.89"),
        NumberCell("V1", "1234567.89", 3),
        NumberCell("W1", "1234567.89", 1),
        NumberCell("Z1", "1234567.89"),
        NumberCell("AA1", "1234567.89", 0),
        NumberCell("AB1", "1234567.89", 6),
        NumberCell("AC1", "123456789", 9),
        NumberCell("AD1", "-1234567.89", 9),
        NumberCell("AE1", "NaN", 1),
        NumberCell("AF1", "Infinity", 1),
        NumberCell("AG1", "-Infinity", 1),
        NumberCell("AH1", "-1.5", 11),
        NumberCell("GS1", "12345.67", 1),
        NumberCell("GT1", "98765.43", 1))
    { RowIndex = 1 };
    var row2 = new Row(NumberCell("J2", "1234567.89", 1))
    {
        RowIndex = 2,
        Hidden = true
    };
    var row3 = new Row(NumberCell("U3", "1234567.89"))
    {
        RowIndex = 3,
        CustomFormat = true,
        StyleIndex = 1
    };
    var row4 = new Row(NumberCell("X4", "1234567.89", 1))
    {
        RowIndex = 4,
        CustomFormat = true,
        StyleIndex = 3
    };
    var row5 = new Row(NumberCell("Y5", "1234567.89"))
    {
        RowIndex = 5,
        CustomFormat = true,
        StyleIndex = 3
    };
    var row5001 = new Row(
        NumberCell("Q5001", "12345.67", 1),
        NumberCell("R5001", "98765.43", 1))
    { RowIndex = 5001 };

    var sheetData = new SheetData(row1, row2, row3, row4, row5, row5001);
    var mergeCells = new MergeCells(
        new MergeCell { Reference = "G1:H1" },
        new MergeCell { Reference = "Q5001:R5001" },
        new MergeCell { Reference = "GS1:GT1" })
    { Count = 3 };
    return new Worksheet(
        new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 8.43D },
        columns,
        sheetData,
        mergeCells);
}

static Worksheet CreateFormulaViewWorksheet()
{
    return new Worksheet(
        new SheetViews(new SheetView { WorkbookViewId = 0, ShowFormulas = true }),
        new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 2D },
        new Columns(Col(1, 2), Col(2, 2)),
        new SheetData(new Row(
            NumberCell("A1", "1234567.89", 1),
            FormulaCell("B1", "1234567.89", 1))
        { RowIndex = 1 }));
}

static Worksheet CreateHiddenWorksheet()
{
    return new Worksheet(
        new SheetFormatProperties { DefaultRowHeight = 15D, DefaultColumnWidth = 2D },
        new Columns(Col(1, 2)),
        new SheetData(new Row(NumberCell("A1", "1234567.89", 1)) { RowIndex = 1 }));
}

static Column Col(uint index, double width, bool hidden = false, uint? styleIndex = null)
{
    var column = new Column
    {
        Min = index,
        Max = index,
        Width = width,
        CustomWidth = true,
        Hidden = hidden
    };
    if (styleIndex.HasValue) column.Style = styleIndex.Value;
    return column;
}

static Cell NumberCell(string reference, string value, uint? styleIndex = null)
{
    var cell = new Cell
    {
        CellReference = reference,
        CellValue = new CellValue(value)
    };
    if (styleIndex.HasValue) cell.StyleIndex = styleIndex.Value;
    return cell;
}

static Cell FormulaCell(string reference, string formula, uint styleIndex) => new()
{
    CellReference = reference,
    StyleIndex = styleIndex,
    CellFormula = new CellFormula(formula)
};

static Cell IsoDateCell(string reference, string value, uint styleIndex) => new()
{
    CellReference = reference,
    StyleIndex = styleIndex,
    DataType = CellValues.Date,
    CellValue = new CellValue(value)
};

static Cell InlineTextCell(string reference, string value) => new()
{
    CellReference = reference,
    StyleIndex = 1,
    DataType = CellValues.InlineString,
    InlineString = new InlineString(new Text(value))
};

static void AssertPaths(IEnumerable<string> expected, IEnumerable<string> actual, string scenario)
{
    var expectedList = expected.OrderBy(path => path, StringComparer.Ordinal).ToArray();
    var actualList = actual.OrderBy(path => path, StringComparer.Ordinal).ToArray();
    if (expectedList.SequenceEqual(actualList, StringComparer.Ordinal)) return;
    throw new InvalidOperationException(
        $"{scenario}: expected [{string.Join(", ", expectedList)}], got [{string.Join(", ", actualList)}]");
}

static void Assert(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}
