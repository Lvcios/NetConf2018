﻿using DataSources;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StyleSheet
{
    class Program
    {
        static void Main(string[] args)
        {
            IEnumerable<Customer> reportData = Report.GetCustomers();

            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create("CustomersReport_Styled.xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart wBookPart = null;
                wBookPart = spreadsheetDoc.AddWorkbookPart();
                wBookPart.Workbook = new Workbook();
                spreadsheetDoc.WorkbookPart.Workbook.Sheets = new Sheets();
                Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                WorksheetPart wSheetPart = wBookPart.AddNewPart<WorksheetPart>();


                WorkbookStylesPart stylesPart = spreadsheetDoc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = StylesheetGenerator.GenerateStyleSheet();
                stylesPart.Stylesheet.Save();


                Columns columns = new Columns();
                columns.Append(new Column { Width = 30, Min = 1, Max = 8 });

                string relationshipId = spreadsheetDoc.WorkbookPart.GetIdOfPart(wSheetPart);
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(wSheetPart),
                    SheetId = sheetId,
                    Name = "Hoja_" + sheetId,
                };
                sheets.Append(sheet);

                SheetData sheetData = new SheetData();
                wSheetPart.Worksheet = new Worksheet(columns, sheetData);

                Row headerRow = new Row();
                headerRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = "Name" } });
                headerRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = "Register Date" } });
                headerRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = "Last Buy" } });
                headerRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = "Product" } });
                headerRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = "Cost" } });
                headerRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = "Quantity" } });
                headerRow.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue { Text = "Total" } });

                sheetData.AppendChild(headerRow);

                foreach (Customer data in reportData)
                {
                    Row contentRow = new Row();
                    contentRow.Append(new Cell { StyleIndex = 3, DataType = CellValues.String, CellValue = new CellValue { Text = data.Name } });
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
