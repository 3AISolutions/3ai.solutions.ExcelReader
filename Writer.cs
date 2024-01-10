using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace _3ai.solutions.ExcelReader;

public static partial class Writer
{
  public static byte[] CreateExcel<T>(this IEnumerable<T> items)
  {
      using var stream = new MemoryStream();
      using (var spreadsheetDocument = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
      {
          var workbookPart = spreadsheetDocument.AddWorkbookPart();
          var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
          stylesPart.Stylesheet = new Stylesheet
          {
              Fonts = new Fonts(new Font()),
              Fills = new Fills(new Fill()),
              Borders = new Borders(new Border()),
              CellStyleFormats = new CellStyleFormats(new CellFormat()),
              CellFormats =
                  new CellFormats(
                      new CellFormat(),
                      new CellFormat
                      {
                          NumberFormatId = 14,
                          ApplyNumberFormat = true
                      })
          };
  
          workbookPart.Workbook = new Workbook();
          var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
          var sheetData = new SheetData();
          worksheetPart.Worksheet = new Worksheet(sheetData);
          var sheets = workbookPart.Workbook.AppendChild(new Sheets());
          var sheet = new Sheet
          {
              Id = workbookPart.GetIdOfPart(worksheetPart),
              SheetId = 1,
              Name = typeof(T).Name
          };
          sheets.Append(sheet);
          var row = new Row();
          var columns = typeof(T).GetProperties();
          foreach (var column in columns)
          {
              var cell = new Cell
              {
                  DataType = CellValues.String,
                  CellValue = new CellValue(column.Name)
              };
              row.AppendChild(cell);
          }
          sheetData.AppendChild(row);
          foreach (var item in items)
          {
              row = new Row();
              foreach (var column in columns)
              {
                  var cell = new Cell();
                  var value = column.GetValue(item);
                  if (value is DateTime time)
                  {
                      cell.CellValue = new CellValue(time);
                      cell.DataType = CellValues.Date;
                      cell.StyleIndex = 1;
                  }
                  else if (value is DateTimeOffset offset)
                  {
                      cell.CellValue = new CellValue(offset);
                      cell.DataType = CellValues.Date;
                      cell.StyleIndex = 1;
                  }
                  else if (value is bool v)
                  {
                      cell.CellValue = new CellValue(v);
                      cell.DataType = CellValues.Boolean;
                  }
                  else if (value is double d)
                  {
                      cell.DataType = CellValues.Number;
                      cell.CellValue = new CellValue(d);
                  }
                  else if (value is decimal dec)
                  {
                      cell.DataType = CellValues.Number;
                      cell.CellValue = new CellValue(dec);
                  }
                  else if (value is int l)
                  {
                      cell.DataType = CellValues.Number;
                      cell.CellValue = new CellValue(l);
                  }
                  else if (value is string s)
                  {
                      cell.DataType = CellValues.String;
                      cell.CellValue = new CellValue(s);
                  }
                  else if (value is null)
                  {
                      cell.CellValue = null;
                  }
                  else
                  {
                      cell.DataType = CellValues.String;
                      cell.CellValue = new CellValue(value?.ToString() ?? "");
                  }
                  row.AppendChild(cell);
              }
              sheetData.AppendChild(row);
          }
      }
      return stream.ToArray();
  }
}
