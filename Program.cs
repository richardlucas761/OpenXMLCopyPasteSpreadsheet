using DocumentFormat.OpenXml.Packaging;

const string ToDoList = "ToDoList.xlsx";
const string InventoryList = "InventoryList.xlsx";
const string ToDoListSheet = "To-do list";
const string InventoryListSheet = "Inventory list";

const string Output1 = "Output1.xlsx";
const string Output2 = "Output2.xlsx";
const string Output3 = "Output3.xlsx";
const string Output4 = "Output4.xlsx";

// ----------------------------------------------------------------------------
// General notes
// ----------------------------------------------------------------------------
// All experiments take place in the "\bin\Debug\net10.0" folder.
// Files are copied to "OutputN.xlsx" to preserve original documents for comparison.
// ----------------------------------------------------------------------------

CleanUp(Output1, Output2, Output3, Output4);
Experiment1(ToDoList, InventoryList, InventoryListSheet, Output1);
Experiment2(ToDoListSheet);
Experiment3(ToDoList, InventoryList, ToDoListSheet, Output3);
Experiment4(ToDoList, ToDoListSheet, Output4);

static void Experiment4(string ToDoList, string ToDoListSheet, string Output4)
{
    // ----------------------------------------------------------------------------
    // Experiment 4 "Output4.xlsx" - Merge two spreadsheets together, more complex spreadsheet
    // with frozen rows, word wrap, merges and an image present.
    // ----------------------------------------------------------------------------

    File.Copy("WithImage.xlsx", Output4, true);

    OpenXMLCopyPasteSpreadsheet.MergeXLSX.MergeXSLX(ToDoList, Output4, [ToDoListSheet, "Assignment setup"]);

    // Note: The same Security Warning "Automatic update of external links has been disabled." is seen in LibreOffice for
    // "Output4.xlsx".

    // Note: after uploading "Output4.xlsx" to Office.com and viewing the spreadsheet in the web "Open in browser" this message
    // was seen, suggesting the file produced is corrupted:
    // "WORKBOOK REPAIRED We temporarily repaired this workbook so that you can open it in Reading View."
    // Technical details below:
    /*
    Repaired Records: Conditional formatting from /xl/worksheets/sheet116.xml part
    Repaired Records: Table from /xl/tables/table112.xml part (Table)
    Repaired Records: Table from /xl/tables/table213.xml part (Table)
    */
}

static void Experiment3(string ToDoList, string InventoryList, string ToDoListSheet, string Output3)
{
    // ----------------------------------------------------------------------------
    // Experiment 3 "Output3.xlsx" - Merge two spreadsheets together.
    // ----------------------------------------------------------------------------

    File.Copy(InventoryList, Output3, true);

    OpenXMLCopyPasteSpreadsheet.MergeXLSX.MergeXSLX(ToDoList, Output3, [ToDoListSheet, "Assignment setup"]);

    // Note: When viewing the created file "Output3.xlsx" in LibreOffice a Security Warning "Automatic update of external links has
    // been disabled." is displayed. This is not seen for "Output1.xlsx" and "Output2.xlsx" for the previous two experiments. The
    // reason for this is not known.

    // Note: after uploading "Output3.xlsx" to Office.com and viewing the spreadsheet in the web "Open in browser" this message
    // was seen, suggesting the file produced is corrupted:
    // "WORKBOOK REPAIRED We temporarily repaired this workbook so that you can open it in Reading View."
    // Technical details below:
    /*
     Repaired Records: Conditional formatting from /xl/worksheets/sheet112.xml part
     Repaired Records: Table from /xl/tables/table112.xml part (Table)
     Repaired Records: Table from /xl/tables/table213.xml part (Table)
     */
}

static void Experiment2(string ToDoListSheet)
{
    // ----------------------------------------------------------------------------
    // Experiment 2 "Output2.xlsx" - Clone an existing sheet in a spreadsheet.
    // ----------------------------------------------------------------------------

    const string experiment2 = Output2;

    File.Copy(ToDoList, experiment2, true);

    using var spreadsheet = SpreadsheetDocument.Open(experiment2, true);

    OpenXMLCopyPasteSpreadsheet.CloningSheet.CloneSheet(spreadsheet, ToDoListSheet, ToDoListSheet + "2");

    // Note: after uploading "Output2.xlsx" to Office.com and viewing the spreadsheet in the web "Open in browser" this message
    // was seen, suggesting the file produced is corrupted:
    // "WORKBOOK REPAIRED We temporarily repaired this workbook so that you can open it in Reading View."
    // Technical details below:
    /*
     Repaired Records: Table from /xl/tables/table113.xml part (Table)
     */
}

static void Experiment1(string ToDoList, string InventoryList, string InventoryListSheet, string Output1)
{
    // ----------------------------------------------------------------------------
    // Experiment 1 "Output1.xlsx" - The "Inventory list" sheet is copied twice from
    // "InventoryList.xlsx" to "Output1.xlsx". Note this overwrites any existing content
    // in "Output1.xlsx" which may not be desirable.
    // ----------------------------------------------------------------------------

    // Any content in "Output1.xlsx" will be overwritten.
    File.Copy(ToDoList, Output1, true);

    // "Inventory list" is created in "Output1.xlsx".
    OpenXMLCopyPasteSpreadsheet.OpenXMLCopySheet.CopySheet(InventoryList, InventoryListSheet,
    InventoryListSheet, Output1);

    // "Inventory list2" is created in "Output1.xlsx" which also deletes any other content in "Output1.xlsx".
    OpenXMLCopyPasteSpreadsheet.OpenXMLCopySheet.CopySheet(InventoryList, InventoryListSheet,
        InventoryListSheet + "2", Output1);

    // Note: after uploading "Output1.xlsx" to Office.com and viewing the spreadsheet in the web "Open in browser" this message
    // was seen, suggesting the file produced is corrupted:
    // "WORKBOOK REPAIRED We temporarily repaired this workbook so that you can open it in Reading View."
    // Technical details below:
    /*
    Removed Records: Table from /xl/tables/table11.xml part (Table)
    Repaired Records: Cell information from /xl/worksheets/sheet11.xml part
    Repaired Records: Column information from /xl/worksheets/sheet11.xml part
    Repaired Records: Conditional formatting from /xl/worksheets/sheet11.xml part
    Repaired Records: Table from /xl/tables/table11.xml part (Table)
    */
}

static void CleanUp(string Output1, string Output2, string Output3, string Output4)
{
    // ----------------------------------------------------------------------------
    // Clean up any old files from previous experiments
    // ----------------------------------------------------------------------------

    if (File.Exists(Output1))
    {
        File.Delete(Output1);
    }

    if (File.Exists(Output2))
    {
        File.Delete(Output2);
    }

    if (File.Exists(Output3))
    {
        File.Delete(Output3);
    }

    if (File.Exists(Output4))
    {
        File.Delete(Output4);
    }
}
