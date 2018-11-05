using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class TableHelper
    {
        public static void AddTable()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("TableSheet");

            // Inset table data
            worksheet.Cells["A1"].Value = "Name";
            worksheet.Cells["A2"].Value = "Orange";
            worksheet.Cells["A3"].Value = "Apple";
            worksheet.Cells["A4"].Value = "Banana";

            worksheet.Cells["B1"].Value = "Price";
            worksheet.Cells["B2"].Value = 12;
            worksheet.Cells["B3"].Value = 16;
            worksheet.Cells["B4"].Value = 9;

            worksheet.Cells["C1"].Value = "Quantity";
            worksheet.Cells["C2"].Value = 50;
            worksheet.Cells["C3"].Value = 64;
            worksheet.Cells["C4"].Value = 72;

            var range = worksheet.Cells["A1:C4"];

            //Create table from data
            var table = worksheet.Tables.Add(range, "Table");
            //Showing additional end row
            table.ShowTotal = true;
            table.Columns[0].TotalsRowLabel = "Total";
            //Set table functions
            table.Columns[1].TotalsRowFunction = Table.RowFunctions.Average;
            table.Columns[2].TotalsRowFunction = Table.RowFunctions.Sum;
            //Choose table style
            table.TableStyle = Table.TableStyles.Light1;

            excelApp.Save("table.xlsx");
        }
    }
}
