using iDiTect.Excel.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class PictureHelper
    {
        public static void InsertPictureToSheet()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("PictureSheet");

            //Insert image to worksheet
            Image pic1 = Image.FromFile("demo-pic.png");

            Picture excelPic = worksheet.Drawings.AddPicture("pic1", pic1);
            //Insert picture to cell "B2" with 0 offset
            excelPic.SetPosition(2, 0, 2, 0);
            //Customize the picture size
            //picture will keey the original size while not calling this method
            excelPic.SetSize(300, 150);


            //Insert picture with hyperlink
            Uri link = new Uri("http://www.iditect.com");
            Picture hyperPic = worksheet.Drawings.AddPicture("pic2", "demo-pic.png", link);
            hyperPic.SetPosition(10, 10, 2, 10);

            excelApp.Save("picture.xlsx");
        }
    }
}
