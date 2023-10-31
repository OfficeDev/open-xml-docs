#nullable disable

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace GetAllWorkheets
{
    class Program
    {
        const string DEMOFILE =
            @"C:\Users\Public\Documents\SampleWorkbook.xlsx";

        static void Main(string[] args)
        {
            var results = GetAllWorksheets(DEMOFILE);
            foreach (Sheet item in results)
            {
                Console.WriteLine(item.Name);
            }
        }

        // Retrieve a List of all the sheets in a workbook.
        // The Sheets class contains a collection of 
        // OpenXmlElement objects, each representing one of 
        // the sheets.
        public static Sheets GetAllWorksheets(string fileName)
        {
            Sheets theSheets = null;

            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                theSheets = wbPart.Workbook.Sheets;
            }
            return theSheets;
        }
    }
}