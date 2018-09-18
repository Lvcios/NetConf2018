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
    class Program
    {
        static void Main(string[] args)
        {
            Hard.CreateExcel();
            Easy.CreateExcel();
        }
    }
}
