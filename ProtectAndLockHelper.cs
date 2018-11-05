using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class ProtectAndLockHelper
    {
        public static void ProtectExcel()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("LockSheet");

            //Insert 10 rows data
            var rnd = new Random();
            for (int row = 1; row <= 10; row++)
            {
                worksheet.Cells[row, 1].Value = row;
                worksheet.Cells[row, 2].Value = string.Format("Row {0}", row);
                worksheet.Cells[row, 3].Value = rnd.Next();
            }

            //Lock for cell
            worksheet.Cells[1, 1].Style.Locked = true;
            //Lock for row
            worksheet.Row(2).Style.Locked = true;
            //Lock for column
            worksheet.Column(2).Style.Locked = true;

            //Simply lock the sheet data, all the cells will be read only.
            worksheet.Protection.IsProtected = true;
            //Password protect, the Protection.IsProtected will be set to true automatically
            //worksheet.Protection.SetPassword("password 1");

            //Unlock the third row
            worksheet.Row(3).Style.Locked = false;
            //Unlock specified cells
            worksheet.Cells[7, 1, 8, 2].Style.Locked = false;

            //Protect whole excel document by password
            excelApp.Save("lock.xlsx", "password 2");

        }
    }
}
