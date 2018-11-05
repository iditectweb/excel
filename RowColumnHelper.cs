using iDiTect.Excel.Licensing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class RowColumnHelper
    {
        public static void RowAndColumn()
        {            
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("RowAndColumn");
         

            //Insert 1 row from the first row
            worksheet.InsertRow(1, 1);
            //Select the first row
            var row = worksheet.Row(1);
            //Set row style
            row.Style.Fill.PatternType = Style.FillStyle.Solid;
            row.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
            //height is in units of pounds (1cm = 28.6 pounds) 
            row.Height = 20;
            

            //Insert 1 column from the sencond column
            worksheet.InsertColumn(2, 1);
            //Select the second column (Column B)
            var column = worksheet.Column(2);
            //Set column style
            column.Style.Fill.PatternType = Style.FillStyle.Solid;
            column.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
            //width is in units of 1/10 inch (2.54 millimeters)
            column.Width = 12;


            //Select C2 as a Range
            var range1 = worksheet.Cells["C2"];
            range1.Value = "this is long text";
            range1.AutoFitColumns();

            //Get a Range from cells C3:E5
            var range2 = worksheet.Cells[3, 3, 5, 5];
            range2.Merge = true;
            range2.Value = "merged cells";
            range2.Style.HorizontalAlignment = Style.HorizontalAlignment.Center;
            range2.Style.VerticalAlignment = Style.VerticalAlignment.Center;

            //Select all cells in worksheet
            var range3 = worksheet.Cells["A: XFD"];
            //Set range style
            range3.Style.Font.Color.SetColor(System.Drawing.Color.Gray); 


            excelApp.Save("row-column.xlsx");
        }



    }
}
