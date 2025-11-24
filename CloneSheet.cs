using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLCopyPasteSpreadsheet
{
    public static class CloningSheet
    {
        // Source - https://stackoverflow.com/a/61139061
        // Posted by Patryk Sładek
        // Retrieved 2025-11-24, License - CC BY-SA 4.0

        static void CloneSheet(SpreadsheetDocument spreadsheetDocument, string sheetName, string clonedSheetName)
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart sourceSheetPart = OpenXMLCopySheet.GetWorkSheetPart(workbookPart, sheetName);
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

            SpreadsheetDocument tempSheet = SpreadsheetDocument.Create(new MemoryStream(), spreadsheetDocument.DocumentType);
            WorkbookPart tempWorkbookPart = tempSheet.AddWorkbookPart();
            WorksheetPart tempWorksheetPart = tempWorkbookPart.AddPart(sourceSheetPart);
            WorksheetPart clonedSheet = workbookPart.AddPart(tempWorksheetPart);

            Sheet copiedSheet = new Sheet();
            copiedSheet.Name = clonedSheetName;
            copiedSheet.Id = workbookPart.GetIdOfPart(clonedSheet);
            copiedSheet.SheetId = (uint)sheets.ChildElements.Count + 1;
            sheets.Append(copiedSheet);
        }

    }
}
