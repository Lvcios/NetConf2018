﻿using DataSources;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template
{
    class Program
    {
        static void Main(string[] args)
        {
            //Get Template and create new file to edit
            System.IO.File.Copy("Template.xlsx", "CustomersReport.xlsx", true);

            //Generate Data
            IEnumerable<Customer> reportData = Report.GetCustomers();

            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open("CustomersReport.xlsx", true))
            {
                WorkbookPart wBookPart = null;
                wBookPart = spreadsheetDoc.WorkbookPart;
                spreadsheetDoc.WorkbookPart.Workbook.Sheets = new Sheets();
                Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>();

                //Get worksheetpart
                WorksheetPart wSheetPart = spreadsheetDoc.WorkbookPart.WorksheetParts.First();

                //Get existing workSheetPart
                WorksheetPart newWorksheetPart = spreadsheetDoc.WorkbookPart.WorksheetParts.First();

                //add Styles
                WorkbookStylesPart stylesPart = spreadsheetDoc.WorkbookPart.WorkbookStylesPart;
                //stylesPart.Stylesheet = Styles.GenerateStyleSheet(); //ya tiene una hoja de estilos
                stylesPart.Stylesheet.Save();

                string relationshipId = spreadsheetDoc.WorkbookPart.GetIdOfPart(newWorksheetPart);

            
                // Give the new worksheet a name.
                Sheet sheet = new Sheet {
                    Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(newWorksheetPart),
                    SheetId = 1,
                    Name = "Customer_Report" };
                sheets.Append(sheet);

                //get existing sheetData
                SheetData sheetData = newWorksheetPart.Worksheet.GetFirstChild<SheetData>();

                foreach (Customer data in reportData)
                {
                    Row contentRow = new Row();
                    contentRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = data.Name } });
                    contentRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = data.RegisterDate } });
                    contentRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = data.LastBuy } });
                    contentRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = data.Item } });
                    contentRow.Append(new Cell { DataType = CellValues.Number, CellValue = new CellValue { Text = data.Quantity.ToString() } });
                    contentRow.Append(new Cell { DataType = CellValues.Number, CellValue = new CellValue { Text = data.ItemCost.ToString() } });
                    contentRow.Append(new Cell { DataType = CellValues.Number, CellValue = new CellValue { Text = string.Format("{0}", data.Quantity * data.ItemCost) } });
                    sheetData.AppendChild(contentRow);
                }
            }

        }
    }
}
