﻿using DataSources;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleExcel
{
    public class Hard
    {
        public static void CreateExcel()
        {
            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Create("CustomerReport_Hard.xlsx", SpreadsheetDocumentType.Workbook))
            {
                OpenXmlWriter oxw;
                List<OpenXmlAttribute> oxa;
                spreadsheetDoc.AddWorkbookPart();
                WorksheetPart wsp = spreadsheetDoc.WorkbookPart.AddNewPart<WorksheetPart>();
                oxw = OpenXmlWriter.Create(wsp);

                oxw.WriteStartElement(new Worksheet());
                
                //columnas
                oxw.WriteStartElement(new Columns());
                oxa = new List<OpenXmlAttribute>();
                oxa.Add(new OpenXmlAttribute("min", null, "1"));
                oxa.Add(new OpenXmlAttribute("max", null, "4"));
                oxa.Add(new OpenXmlAttribute("width", null, "25"));
                oxw.WriteStartElement(new Column(), oxa);
                oxw.WriteEndElement();
               
                oxa = new List<OpenXmlAttribute>();
                oxa.Add(new OpenXmlAttribute("min", null, "6"));
                oxa.Add(new OpenXmlAttribute("max", null, "6"));
                oxa.Add(new OpenXmlAttribute("width", null, "40"));
                oxw.WriteStartElement(new Column(), oxa);
                oxw.WriteEndElement();

                oxw.WriteEndElement();
                
                oxw.WriteStartElement(new SheetData());
                oxw.WriteStartElement(new Row(), new List<OpenXmlAttribute>());
                oxa = new List<OpenXmlAttribute>();
                oxa.Add(new OpenXmlAttribute("t", null, "str"));

                
                oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue("Name")); oxw.WriteEndElement();
                oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue("Register Date")); oxw.WriteEndElement();
                oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue("Last Buy")); oxw.WriteEndElement();
                oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue("Product")); oxw.WriteEndElement();
                oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue("Cost")); oxw.WriteEndElement();
                oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue("Quantity")); oxw.WriteEndElement();
                oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue("Total")); oxw.WriteEndElement();

                oxw.WriteEndElement();
                
                foreach(Customer customer in Report.GetCustomers())
                {
                    oxw.WriteStartElement(new Row(), new List<OpenXmlAttribute>());
                    oxa = new List<OpenXmlAttribute>();
                    oxa.Add(new OpenXmlAttribute("t", null, "str"));

                    oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue(customer.Name)); oxw.WriteEndElement();
                    oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue(customer.RegisterDate)); oxw.WriteEndElement();
                    oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue(customer.LastBuy)); oxw.WriteEndElement();
                    oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue(customer.Item)); oxw.WriteEndElement();
                    oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue(customer.ItemCost.ToString())); oxw.WriteEndElement();
                    oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue(customer.Quantity.ToString())); oxw.WriteEndElement();
                    oxw.WriteStartElement(new Cell(), oxa); oxw.WriteElement(new CellValue((customer.Quantity * customer.ItemCost).ToString())); oxw.WriteEndElement();
                    oxw.WriteEndElement();
                }

                oxw.WriteEndElement();
                oxw.WriteEndElement();
                oxw.Close();

                oxw = OpenXmlWriter.Create(spreadsheetDoc.WorkbookPart);
                oxw.WriteStartElement(new Workbook());
                oxw.WriteStartElement(new Sheets());

                oxw.WriteElement(new Sheet()
                {
                    Name = "Sheet1",
                    SheetId = 1,
                    Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(wsp)
                });

                
                oxw.WriteEndElement();
                
                oxw.WriteEndElement();
                oxw.Close();

                spreadsheetDoc.Close();
            }
        }
    }
}
