using iDiTect.Excel.Sparkline;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class SparklineHelper
    {
        public static void AddSparkline()
        {
            var dt = GetDataTable();

            var excelApp = new Application();
            
            var worksheet = excelApp.Workbook.Worksheets.Add("Sparkline");

            //Bind data table to cells
            worksheet.Cells["A1"].LoadFromDataTable(dt, true);
            
            //Add column type sparkline
            worksheet.Cells["N1"].Value = "Column";
            var sparklineCol = worksheet.SparklineGroups.Add(SparklineType.Column, worksheet.Cells["N2:N7"], worksheet.Cells["B2:M7"]);
            //Highlight the highest data
            sparklineCol.High = true;
            sparklineCol.ColorHigh.SetColor(System.Drawing.Color.Red);
            //Highlight the lowest data
            sparklineCol.Low = true;
            sparklineCol.ColorLow.SetColor(System.Drawing.Color.Green);

            //Add line type sparkline
            worksheet.Cells["O1"].Value = "Line";
            var sparklineLine = worksheet.SparklineGroups.Add(SparklineType.Line, worksheet.Cells["O2:O7"], worksheet.Cells["B2:M7"]);
           
            excelApp.Save("Sparkline.xlsx");
        }

        private static DataTable GetDataTable()
        {
            string[] products = { "Orange", "Apple", "Banana", "Pear", "Grape", "Watermelon" };
            DataTable dt = new DataTable("Fruit Sales");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Jan", typeof(long));
            dt.Columns.Add("Feb", typeof(long));
            dt.Columns.Add("Mar", typeof(long));
            dt.Columns.Add("Apr", typeof(long));
            dt.Columns.Add("May", typeof(long));
            dt.Columns.Add("Jun", typeof(long));
            dt.Columns.Add("Jul", typeof(long));
            dt.Columns.Add("Aug", typeof(long));
            dt.Columns.Add("Sep", typeof(long));
            dt.Columns.Add("Oct", typeof(long));
            dt.Columns.Add("Nov", typeof(long));
            dt.Columns.Add("Dec", typeof(long));

            Random r = new Random();
            for (int row = 0; row < 6; row++)
            {
                var dr = dt.NewRow();
                dr[0] = products[row];

                for (int col = 1; col < 13; col++)
                {
                    dr[col] = r.Next(10, 100);
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }
    }
}
