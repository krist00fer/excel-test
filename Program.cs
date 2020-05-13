using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace excel_test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            CreateSpreadsheetWorkbook("test-doc.xlsx");

        }

        public static void CreateSpreadsheetWorkbook(string filepath)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {

                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
                sheets.Append(sheet);

                var cellA1 = InsertCellInWorksheet("A", 1, worksheetPart);
                cellA1.CellValue = new CellValue("Hello World!");
                cellA1.DataType = new EnumValue<CellValues>(CellValues.String);

                var cellB1 = InsertCellInWorksheet("B", 1, worksheetPart);
                cellB1.CellValue = new CellValue("3.14159");
                cellB1.DataType = new EnumValue<CellValues>(CellValues.Number);

                workbookpart.Workbook.Save();

            }

            
        }

                    // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
            // If the cell already exists, returns it. 
            private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
            {
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                string cellReference = columnName + rowIndex;

                // If the worksheet does not contain a row with the specified row index, insert one.
                Row row;
                if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
                {
                    row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
                }
                else
                {
                    row = new Row() { RowIndex = rowIndex };
                    sheetData.Append(row);
                }

                // If there is not a cell with the specified column name, insert one.  
                if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
                {
                    return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
                }
                else
                {
                    // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                    Cell refCell = null;
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }

                    Cell newCell = new Cell() { CellReference = cellReference };
                    row.InsertBefore(newCell, refCell);

                    worksheet.Save();
                    return newCell;
                }
            }
    }
}
