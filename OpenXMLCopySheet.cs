// Source - https://stackoverflow.com/q/50856596
// Posted by Mike F, modified by community. See post 'Timeline' for change history
// Retrieved 2025-11-24, License - CC BY-SA 4.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLCopyPasteSpreadsheet
{
    static public class OpenXMLCopySheet
    {
        private const string CopiedTable = "CopiedTable";

        private static int tableId;

        public static void CopySheet(string filename, string sheetName, string clonedSheetName, string destFileName)
        {
            //Open workbook
            using var mySpreadsheet = SpreadsheetDocument.Open(filename, true);

            WorkbookPart workbookPart = mySpreadsheet.WorkbookPart;

            //Get the source sheet to be copied
            WorksheetPart sourceSheetPart = GetWorkSheetPart(workbookPart, sheetName);
            SharedStringTablePart sharedStringTable = workbookPart.SharedStringTablePart;

            //Take advantage of AddPart for deep cloning
            using var newXLSFile = SpreadsheetDocument.Create(destFileName, SpreadsheetDocumentType.Workbook);

            WorkbookPart newWorkbookPart = newXLSFile.AddWorkbookPart();
            SharedStringTablePart newSharedStringTable = newWorkbookPart.AddPart<SharedStringTablePart>(sharedStringTable);
            WorksheetPart newWorksheetPart = newWorkbookPart.AddPart<WorksheetPart>(sourceSheetPart);

            //Table definition parts are somewhat special and need unique ids...so let's make an id based on count
            int numTableDefParts = sourceSheetPart.GetPartsOfType<TableDefinitionPart>().Count();

            tableId = numTableDefParts;

            //Clean up table definition parts (tables need unique ids)
            if (numTableDefParts != 0)
                FixupTableParts(newWorksheetPart, numTableDefParts);
            //There should only be one sheet that has focus
            CleanView(newWorksheetPart);

            var fileVersion = new FileVersion { ApplicationName = "Microsoft Office Excel" };

            //Worksheet ws = newWorksheetPart.Worksheet;
            Workbook wb = new Workbook();
            wb.Append(fileVersion);

            //Add new sheet to main workbook part
            Sheets sheets = null;
            //int sheetCount = wb.Sheets.Count();
            if (wb.Sheets != null)
            { sheets = wb.GetFirstChild<Sheets>(); }
            else
            { sheets = new Sheets(); }

            Sheet copiedSheet = new Sheet
            {
                Name = clonedSheetName,
                Id = newWorkbookPart.GetIdOfPart(newWorksheetPart)
            };
            if (wb.Sheets != null)
            { copiedSheet.SheetId = (uint)sheets.ChildElements.Count + 1; }
            else { copiedSheet.SheetId = 1; }

            sheets.Append(copiedSheet);
            newWorksheetPart.Worksheet.Save();

            wb.Append(sheets);

            //Save Changes
            newWorkbookPart.Workbook = wb;
            wb.Save();
        }

        public static void CleanView(WorksheetPart worksheetPart)
        {
            //There can only be one sheet that has focus
            var views = worksheetPart.Worksheet.GetFirstChild<SheetViews>();

            if (views != null)
            {
                views.Remove();
                worksheetPart.Worksheet.Save();
            }
        }

        public static void FixupTableParts(WorksheetPart worksheetPart, int numTableDefParts)
        {
            //Every table needs a unique id and name
            foreach (TableDefinitionPart tableDefPart in worksheetPart.TableDefinitionParts)
            {
                tableId++;
                tableDefPart.Table.Id = (uint)tableId;
                tableDefPart.Table.DisplayName = CopiedTable + tableId;
                tableDefPart.Table.Name = CopiedTable + tableId;
                tableDefPart.Table.Save();
            }
        }

        public static WorksheetPart GetWorkSheetPart(WorkbookPart workbookPart, string sheetName)
        {
            //Get the relationship id of the sheetname
            var relId = workbookPart.Workbook.Descendants<Sheet>().First(s => s.Name.Value.Equals(sheetName,
                StringComparison.Ordinal)).Id;

            return (WorksheetPart)workbookPart.GetPartById(relId);
        }
    }
}
