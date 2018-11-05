using iDiTect.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class HyperlinkHelper
    {
        public static void AddHyperlinkToSheet()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("HyperlinkSheet");
                       
            //Insert a text hyperlink to external site
            worksheet.Cells["A1"].Hyperlink = new HyperLink("http://www.iditect.com");
            worksheet.Cells["A1"].Value = "go to external site";
            //Set hyperlink style
            worksheet.Cells["A1"].Style.Font.UnderLine = true;
            worksheet.Cells["A1"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
            worksheet.Cells["A1"].AutoFitColumns();


            //Insert a text hyperlink to another sheet
            var worksheet2 = excelApp.Workbook.Worksheets.Add("sheet2");
            worksheet.Cells["A2"].Hyperlink = new HyperLink("sheet2!A1", "go to sheet2");


            //Insert a image hyperlink
            Uri link = new Uri("http://www.iditect.com");
            Picture hyperPic = worksheet.Drawings.AddPicture("picLink", "demo-pic.png", link);
            hyperPic.SetPosition(0, 50);

            excelApp.Save("hyperlink.xlsx");
        }
    }
}
