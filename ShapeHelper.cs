using iDiTect.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class ShapeHelper
    {
        public static void AddShapeToSheet()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ShapeSheet");

            //Add a rectangle shape
            var shape = worksheet.Drawings.AddShape("rectangle", ShapeStyle.Rect);
            //Set shape location and size
            shape.SetPosition(2, 10, 2, 10);
            shape.SetSize(200, 100);

            shape.Text = "First line in shape.\n\rSencond line.";
            //Customize shape fill style
            shape.Fill.Style = FillStyle.SolidFill;
            shape.Fill.Color = Color.LightBlue;
            shape.Fill.Transparancy = 30;
            //Change border style
            shape.Border.Fill.Style = FillStyle.SolidFill;
            shape.Border.LineStyle = LineStyle.Dash;
            shape.Border.Width = 1;
            shape.Border.Fill.Color = Color.Black;
            shape.Border.LineCap = LineCap.Round;
            //Set text alignment
            shape.TextAnchoring = TextAnchoringType.Top;
            shape.TextVertical = TextVerticalType.Horizontal;
            shape.TextAnchoringControl = true;

            excelApp.Save("shape.xlsx");
        }
    }
}
