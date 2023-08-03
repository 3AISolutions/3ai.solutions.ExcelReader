using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace _3ai.solutions.ExcelReader;

public static partial class Reader
{
    public static IEnumerable<List<KeyValuePair<string, string>>> ReadSingleExcelSheet(this Stream stream)
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);

        if (spreadsheetDocument?.WorkbookPart?.Workbook.Sheets is not null)
        {
            var sharedStringTable = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().SingleOrDefault()?.SharedStringTable;
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Single();
            if (sheet.Id is not null && !string.IsNullOrEmpty(sheet.Id.Value))
            {
                if (spreadsheetDocument.WorkbookPart.TryGetPartById(sheet.Id.Value, out var openXmlPart) && openXmlPart is WorksheetPart worksheetPart)
                    return worksheetPart.Worksheet.GetWorksheetData(sharedStringTable).SelectMany(x => x).Select(x => x.ToList());
            }
        }
        return Enumerable.Empty<List<KeyValuePair<string, string>>>();
    }

    public static List<ExcelData> ReadDataFromExcel(this Stream stream)
    {
        var data = new List<ExcelData>();
        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        if (spreadsheetDocument?.WorkbookPart?.Workbook.Sheets is not null)
        {
            var sharedStringTable = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().SingleOrDefault()?.SharedStringTable;
            foreach (var sheet in spreadsheetDocument.WorkbookPart.Workbook.Sheets.Cast<Sheet>())
                if (sheet.Id is not null && !string.IsNullOrEmpty(sheet.Id.Value))
                    if (spreadsheetDocument.WorkbookPart.TryGetPartById(sheet.Id.Value, out var openXmlPart) && openXmlPart is WorksheetPart worksheetPart)
                        foreach (var sheetData in worksheetPart.Worksheet.Elements<SheetData>())
                            data.AddRange(sheetData.GetSheetData(sharedStringTable, sheet.Name?.Value ?? "").SelectMany(x => x));
        }
        return data;
    }

    private static IEnumerable<IEnumerable<IEnumerable<KeyValuePair<string, string>>>> GetWorksheetData(this Worksheet worksheet, SharedStringTable? sharedStringTable)
    {
        foreach (var sheetData in worksheet.Elements<SheetData>())
        {
            yield return sheetData.GetKeyValuePairs(sharedStringTable);
        }
    }

    private static IEnumerable<IEnumerable<KeyValuePair<string, string>>> GetKeyValuePairs(this SheetData sheetData, SharedStringTable? sharedStringTable)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            yield return row.GetKeyValuePairs(sharedStringTable);
        }
    }

    private static IEnumerable<KeyValuePair<string, string>> GetKeyValuePairs(this Row row, SharedStringTable? sharedStringTable)
    {
        foreach (var c in row.Elements<Cell>())
        {
            yield return new(c.GetColumnName(), c.GetCellValue(sharedStringTable));
        }
    }

    private static IEnumerable<IEnumerable<ExcelData>> GetSheetData(this SheetData sheetData, SharedStringTable? sharedStringTable, string sheet)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            yield return row.GetRowData(sharedStringTable, sheet);
        }
    }

    private static IEnumerable<ExcelData> GetRowData(this Row row, SharedStringTable? sharedStringTable, string sheet)
    {
        foreach (var c in row.Elements<Cell>())
        {
            yield return new(sheet, c.GetColumnName(), c.GetRowIndex(), c.GetCellFormula(), c.GetCellValue(sharedStringTable));
        }
    }

    [GeneratedRegex("[A-Za-z]+")]
    private static partial Regex CnRegex();
    private static string GetColumnName(this Cell cell)
    {
        Match match = CnRegex().Match(cell.GetCellReference());
        return match.Value;
    }

    [GeneratedRegex("\\d+")]
    private static partial Regex RiRegex();
    private static uint GetRowIndex(this Cell cell)
    {
        Match match = RiRegex().Match(cell.GetCellReference());
        return uint.Parse(match.Value);
    }

    private static string GetCellReference(this Cell cell)
    {
        return cell.CellReference?.Value ?? "";
    }

    private static string GetCellFormula(this Cell cell)
    {
        return cell.CellFormula?.Text ?? "";
    }

    private static string GetCellValue(this Cell cell, SharedStringTable? sharedStringTable)
    {
        var cellValue = cell.CellValue?.Text ?? "";
        if (cell.DataType is not null)
        {
            switch (cell.DataType.Value)
            {
                case CellValues.Date:
                    cellValue = DateTime.FromOADate(double.Parse(cellValue)).ToString();
                    break;
                case CellValues.SharedString:
                    cellValue = sharedStringTable!.ElementAt(int.Parse(cellValue)).InnerText;
                    break;
                case CellValues.Boolean:
                    if (cellValue.Equals("0"))
                        cellValue = "FALSE";
                    else
                        cellValue = "TRUE";
                    break;
                case CellValues.InlineString:
                    cellValue = cell.InnerText;
                    break;
            }
        }
        return cellValue;
    }

    public static T ReadValue<T>(this string input) where T : struct
    {
        switch (Type.GetTypeCode(typeof(T)))
        {
            case TypeCode.DateTime:
                if (double.TryParse(input, out double oaDate))
                    return (T)(object)DateTime.FromOADate(oaDate);
                break;

            case TypeCode.Int32:
                if (double.TryParse(input, out double parsedNumber) && parsedNumber % 1 == 0)
                    return (T)(object)(int)parsedNumber;
                break;
        }

        throw new FormatException($"{input} could not be converted to {typeof(T).Name}");
    }

    public static T? ReadNullableValue<T>(this string input) where T : struct
    {
        if (string.IsNullOrEmpty(input))
            return null;

        return input.ReadValue<T>();
    }
}
