using iDiTect.Excel.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class ChartHelper
    {
        public static void AddChart()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ChartSheet");

            // Inset chart data
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

            // Set x-series and y-series data from range of cells
            Range xSeries = worksheet.Cells["A2:A4"];
            Range ySeries1 = worksheet.Cells["B2:B4"];
            Range ySeries2 = worksheet.Cells["C2:C4"];


            // Create a Column graph
            var colChart = worksheet.Drawings.AddChart("chart1", ChartType.ColumnClustered);
            // Set chart location and size
            colChart.SetPosition(0, 200);
            colChart.SetSize(400, 300);            
            //Add series
            var series1 = colChart.Series.Add(xSeries, ySeries1);
            series1.Header = "Price";
            var series2 = colChart.Series.Add(xSeries, ySeries2);
            series2.Header = "Quantity";
            // Set chart title
            colChart.Title.Text = "Sample Column Chart";
            //Set chart style
            colChart.Style = ChartStyle.Style1;


            //Add a Line chart
            var lineChart = worksheet.Drawings.AddChart("chart2", ChartType.Line);
            // Set chart location and size
            lineChart.SetPosition(0, 500);
            lineChart.SetSize(400, 300);
            //Add series
            var series3 = lineChart.Series.Add(xSeries, ySeries1);
            series3.Header = "Price";            
            //Set the max and min value for the Y axis
            lineChart.YAxis.MaxValue = 18;
            lineChart.YAxis.MinValue = 8;
            //Set chart style manually
            lineChart.Fill.Color = System.Drawing.Color.DarkCyan;
            lineChart.Border.Width = 3;
            lineChart.Border.Fill.Color = System.Drawing.Color.Black;
            lineChart.PlotArea.Fill.Color = System.Drawing.Color.Yellow;
            series3.Fill.Color = System.Drawing.Color.Blue;
            series3.Border.Width = 2;
            series3.Border.Fill.Color = System.Drawing.Color.Green;


            //Add a 3D Pie chart
            var pieChart = worksheet.Drawings.AddChart("chart3", ChartType.Pie3D);
            // Set chart location and size
            pieChart.SetPosition(0, 800);
            pieChart.SetSize(400, 300);
            //Add series
            var series4 = pieChart.Series.Add(xSeries, ySeries1);
            //3d Settings
            pieChart.View3D.RotX = 45;
            pieChart.View3D.Perspective = 10;


            excelApp.Save("chart.xlsx");
        }

        public static void AddCombinationChart()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ChartSheet");

            // Inset chart data
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

            // Set x-series and y-series data from range of cells
            Range xSeries = worksheet.Cells["A2:A4"];
            Range ySeries1 = worksheet.Cells["B2:B4"];
            Range ySeries2 = worksheet.Cells["C2:C4"];

            //Add a chart with two charttypes (Column and Line) and a secondary axis
            var chart = worksheet.Drawings.AddChart("multi-type chart", ChartType.ColumnClustered);
            // Set chart location and size
            chart.SetPosition(300, 0);
            chart.SetSize(400, 300);
            //Add series
            var series1 = chart.Series.Add(xSeries, ySeries1);
            series1.Header = "Price";

            //Add a Line series
            var lineType = chart.PlotArea.ChartTypes.Add(ChartType.Line);
            //By default the secondary XAxis is not visible
            lineType.UseSecondaryAxis = true;
            lineType.XAxis.Deleted = false;
            lineType.XAxis.TickLabelPosition = TickLabelPosition.High;
            var series2 = lineType.Series.Add(xSeries, ySeries2);
            series2.Header = "Quantity";


            // Set chart title
            chart.Title.Text = "Multi-type Chart";
            //Set chart style
            chart.Style = ChartStyle.Style1;

            excelApp.Save("multi-type-chart.xlsx");
        }
    }
}
