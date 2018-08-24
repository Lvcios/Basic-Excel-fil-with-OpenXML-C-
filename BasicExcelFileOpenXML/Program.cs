using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace BasicExcelFileOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            //Get Template and create new file to edit
            System.IO.File.Copy("Templates\\MediumTemplate.xlsx", "CustomersReport.xlsx", true);

            //Generate Data
            List<CustomersReport> reportData = new CustomersReport().GetRandomData();

            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open("CustomersReport.xlsx", true))
            {
                WorkbookPart wBookPart = null;
                wBookPart = spreadsheetDoc.WorkbookPart;
                wBookPart.Workbook = new Workbook();
                spreadsheetDoc.WorkbookPart.Workbook.Sheets = new Sheets();
                Sheets sheets = spreadsheetDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>();

                //Get worksheetpart
                WorksheetPart wSheetPart = spreadsheetDoc.WorkbookPart.WorksheetParts.First();

                //Get existing workSheetPart
                WorksheetPart newWorksheetPart = spreadsheetDoc.WorkbookPart.WorksheetParts.First();

                //add Styles
                WorkbookStylesPart stylesPart = spreadsheetDoc.WorkbookPart.WorkbookStylesPart;
                //stylesPart.Stylesheet = Styles.GenerateStyleSheet();
                stylesPart.Stylesheet.Save();

                string relationshipId = spreadsheetDoc.WorkbookPart.GetIdOfPart(newWorksheetPart);

                // Get a unique ID for the new worksheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                // Give the new worksheet a name.
                Sheet sheet = new Sheet() { Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(newWorksheetPart), SheetId = sheetId, Name = "Customer_Report" + sheetId };
                sheets.Append(sheet);

                //get existing sheetData
                SheetData sheetData = newWorksheetPart.Worksheet.GetFirstChild<SheetData>();

                foreach (CustomersReport data in reportData)
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
