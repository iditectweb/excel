using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class CreateAndLoadHelper
    {
        public static void CreateNewExcelFile()
        {
            var excelApp = new Application();

            //Create new worksheet by name
            var worksheet = excelApp.Workbook.Worksheets.Add("NewCreateSample");

            //Input cell value by row and column coordinate
            worksheet.Cells[1, 1].Value = "Number";
            worksheet.Cells[1, 2].Value = "String";
            worksheet.Cells[1, 3].Value = "Decimal";
            worksheet.Cells[1, 4].Value = "Time";
            worksheet.Cells[1, 5].Value = "DateTime";

            //Input cell value by range name
            worksheet.Cells["A2"].Value = 12001;
            worksheet.Cells["B2"].Value = "Nails";
            worksheet.Cells["C2"].Value = 37.99;
            worksheet.Cells["D2"].Value = DateTime.Now.ToString("hh:MM:ss");
            worksheet.Cells["E2"].Value = DateTime.Now.ToString("yyyy-mm-dd");

            //Modify the cell format
            worksheet.Cells["A2"].Style.Numberformat.Format = "#,##0";
            worksheet.Cells["C2"].Style.Numberformat.Format = "#,##0.00";
            worksheet.Cells["D2"].Style.Numberformat.Format = "hh:MM:ss";
            worksheet.Cells["E2"].Style.Numberformat.Format = "yyyy-mm-dd";

            //Set core property values
            excelApp.Workbook.Properties.Title = "Sample";
            excelApp.Workbook.Properties.Author = "iDiTect";
            excelApp.Workbook.Properties.Subject = "iDiTect.Excel Sample";
            excelApp.Workbook.Properties.Keywords = "iDiTect.Excel";
            excelApp.Workbook.Properties.Category = "iDiTect.Excel";
            excelApp.Workbook.Properties.Company = "iDiTect.com";
           
            //Set custom property values
            excelApp.Workbook.Properties.SetCustomPropertyValue("CustomProperty1", "Custom Value1");

            excelApp.Save("CreateNew.xlsx");
        }

        public static void LoadFromTemplate()
        {
            string template = "template.xlsx";

            var excelApp = new Application(template);

            //Read all data in worksheet by id
            var templateSheet = excelApp.Workbook.Worksheets[0];

            //Load data from cell by row and column coordinate
            var A1 = templateSheet.Cells[1, 1].Value;
            Console.WriteLine(A1.ToString());

            //Load data from cell by range name
            var C2 = templateSheet.Cells["C2"].Value;
            Console.WriteLine(C2.ToString());

            //After load the existed excel file, you can modify cell value, style, formula, add chart, table...

            //Update cell value or add value to new cell
            templateSheet.Cells["A1"].Value = "update";
            templateSheet.Cells["A10"].Value = "new";
            //... and any other editing

            //Or create a new worksheet
            Worksheet worksheet = excelApp.Workbook.Worksheets.Add("NewTab");
            //do anything you want

            excelApp.Save("modified.xlsx");
        }

       

    }
}
