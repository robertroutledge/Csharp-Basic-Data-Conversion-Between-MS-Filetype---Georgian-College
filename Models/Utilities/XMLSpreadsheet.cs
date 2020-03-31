using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.Globalization;

namespace RoutledgeAssignmentFour.Models.Utilities
{
    public class XMLSpreadsheet
    {
        public static void CreateSpreadsheetWorkbook(String filepath, List<Student> students)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet1 = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Hello"
            };
            sheets.Append(sheet1);

            ////taken from insert cell - try to get it to enter the data in cell A2            
            SharedStringTablePart shareStringPart;
            shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            //Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem("Hello my name is Robert", shareStringPart);
            //// Insert a new worksheet.
           // Insert cell A1 into the new worksheet.
            Cell cell1 = InsertCellInWorksheet("A", 2, worksheetPart);
            //// Set the value of cell.
            cell1.CellValue = new CellValue(index.ToString());
            cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);


            WorksheetPart worksheetPart2 = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart2.Worksheet = new Worksheet(new SheetData());
            Sheet sheet2 = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart2),
                SheetId = 2,
                Name = "studentlist"
            };
            sheets.Append(sheet2);
                                              
            int indexa = InsertSharedStringItem("UniqueID", shareStringPart);
            Cell cella = InsertCellInWorksheet("a", 1, worksheetPart2);
            cella.CellValue = new CellValue(indexa.ToString());
            cella.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int indexb = InsertSharedStringItem("StudentID", shareStringPart);
            Cell cellb = InsertCellInWorksheet("b", 1, worksheetPart2);
            cellb.CellValue = new CellValue(indexb.ToString());
            cellb.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int indexc = InsertSharedStringItem("FirstName", shareStringPart);
            Cell cellc = InsertCellInWorksheet("c", 1, worksheetPart2);
            cellc.CellValue = new CellValue(indexc.ToString());
            cellc.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int indexd = InsertSharedStringItem("LastName", shareStringPart);
            Cell celld = InsertCellInWorksheet("d", 1, worksheetPart2);
            celld.CellValue = new CellValue(indexd.ToString());
            celld.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int indexe = InsertSharedStringItem("DateofBirth", shareStringPart);
            Cell celle = InsertCellInWorksheet("e", 1, worksheetPart2);
            celle.CellValue = new CellValue(indexe.ToString());
            celle.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int indexf = InsertSharedStringItem("IsMe", shareStringPart);
            Cell cellf = InsertCellInWorksheet("f", 1, worksheetPart2);
            cellf.CellValue = new CellValue(indexf.ToString());
            cellf.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            int indexg = InsertSharedStringItem("Age", shareStringPart);
            Cell cellg = InsertCellInWorksheet("g", 1, worksheetPart2);
            cellg.CellValue = new CellValue(indexg.ToString());
            cellg.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            
            ///if this works, just add the rest of the student columns
            ///look for ways to change the datatype
            
            uint rowcount = 2;
            foreach (var student in students)
            {
                char colcount = 'a';
                int index2 = InsertSharedStringItem(student.UniqueID.ToString(), shareStringPart);
                Cell cell2 = InsertCellInWorksheet(colcount.ToString(), rowcount, worksheetPart2);
                cell2.CellValue = new CellValue(index2.ToString());
                cell2.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                colcount = Convert.ToChar(colcount + Convert.ToChar(1));


                int index3 = InsertSharedStringItem(student.StudentId.ToString(), shareStringPart);
                Cell cell3 = InsertCellInWorksheet(colcount.ToString(), rowcount, worksheetPart2);
                cell3.CellValue = new CellValue(index3.ToString());
                cell3.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                colcount = Convert.ToChar(colcount + Convert.ToChar(1));

                int index4 = InsertSharedStringItem(student.FirstName, shareStringPart);
                Cell cell4 = InsertCellInWorksheet(colcount.ToString(), rowcount, worksheetPart2);
                cell4.CellValue = new CellValue(index4.ToString());
                cell4.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                colcount = Convert.ToChar(colcount + Convert.ToChar(1));

                
                int index5 = InsertSharedStringItem(student.LastName, shareStringPart);
                Cell cell5 = InsertCellInWorksheet(colcount.ToString(), rowcount, worksheetPart2);
                cell5.CellValue = new CellValue(index5.ToString());
                cell5.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                colcount = Convert.ToChar(colcount + Convert.ToChar(1));

                string DatestrValue = student.DateOfBirthDT.ToOADate().ToString(CultureInfo.InvariantCulture);
                int index6 = InsertSharedStringItem(DatestrValue, shareStringPart);
                Cell cell6 = InsertCellInWorksheet(colcount.ToString(), rowcount, worksheetPart2);
                cell6.CellValue = new CellValue(index6.ToString());
                cell6.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                colcount = Convert.ToChar(colcount + Convert.ToChar(1));               

                int index7 = InsertSharedStringItem(student.IsMe.ToString(), shareStringPart);
                Cell cell7 = InsertCellInWorksheet(colcount.ToString(), rowcount, worksheetPart2);
                cell7.CellValue = new CellValue(index7.ToString());
                cell7.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                colcount = Convert.ToChar(colcount + Convert.ToChar(1));

                int index8 = InsertSharedStringItem(student.Age.ToString(), shareStringPart);
                Cell cell8 = InsertCellInWorksheet(colcount.ToString(), rowcount, worksheetPart2);
                cell8.CellValue = new CellValue(index8.ToString());
                cell8.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                
                rowcount +=1;                
            }

            // Save and Close the document.
            workbookpart.Workbook.Save();
            spreadsheetDocument.Close();
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

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

        // Given a WorkbookPart, inserts a new worksheet.

        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }


        /// this inserts a new worksheet, need to find a way to have it edit existing. Need one function for create a new sheet, and one for edit existing.
        public static void InsertText(string docName, string text, uint rownum, string colletter)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                // Insert the text into the SharedStringTablePart.
                int index = InsertSharedStringItem(text, shareStringPart);

                // Insert a new worksheet.
                WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);

                // Insert cell A1 into the new worksheet.
                Cell cell = InsertCellInWorksheet(colletter, rownum, worksheetPart);

                // Set the value of cell A1.
                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
                spreadSheet.Close();

            }
        }


    }
}

