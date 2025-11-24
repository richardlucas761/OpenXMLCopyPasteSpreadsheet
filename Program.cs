using DocumentFormat.OpenXml.Packaging;

const string ToDoList = "ToDoList.xlsx";
const string InventoryList = "InventoryList.xlsx";
const string ToDoListSheet = "To-do list";
const string InventoryListSheet = "Inventory list";

const string Output1 = "Output1.xlsx";
const string Output2 = "Output2.xlsx";
const string Output3 = "Output3.xlsx";

// ----------------------------------------------------------------------------
// General notes
// ----------------------------------------------------------------------------
// All experiments take place in the "\bin\Debug\net10.0" folder.
// Files are copied to "OutputN.xlsx" to preserve original documents for comparison.
// ----------------------------------------------------------------------------

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

// ----------------------------------------------------------------------------
// Experiment 1 - The "Inventory list" sheet is copied twice from
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

// ----------------------------------------------------------------------------
// Experiment 2 - Clone an existing sheet in a spreadsheet.
// ----------------------------------------------------------------------------

const string experiment2 = Output2;

File.Copy(ToDoList, experiment2, true);

using var spreadsheet = SpreadsheetDocument.Open(experiment2, true);

OpenXMLCopyPasteSpreadsheet.CloningSheet.CloneSheet(spreadsheet, ToDoListSheet, ToDoListSheet + "2");

// ----------------------------------------------------------------------------
// Experiment 3 - Merge two spreadsheets together.
// ----------------------------------------------------------------------------

File.Copy(InventoryList, Output3, true);

OpenXMLCopyPasteSpreadsheet.MergeXLSX.MergeXSLX(ToDoList, Output3, [ToDoListSheet, "Assignment setup"]);

// Note: When viewing the created file "Output3.xlsx" in LibreOffice a Security Warning "Automatic update of external links has
// been disabled." is displayed. This is not seen for "Output1.xlsx" and "Output2.xlsx" for the previous two experiments. The
// reason for this is not known.
