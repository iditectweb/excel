using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class FormulaHelper
    {
        public static void AddFormula()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("FormulaSheet");
            
            worksheet.Cells["A1"].Value = 1;
            worksheet.Cells["A2"].Value = 2;
            worksheet.Cells["A3"].Value = 3;

            //Add SUM formula
            worksheet.Cells["A5"].Formula = "SUM(A1:A3)";
            worksheet.Cells["A6"].Formula = "1 + 2 + 3";

            //Calculate the formula for cell
            worksheet.Cells["A6"].Calculate();
            //Calculate the formula for whole worksheet
            worksheet.Calculate();

            //Set B2 = A1
            worksheet.Cells["B2"].FormulaR1C1 = "R[-1]C[-1]";
            //Set B3 = A3
            worksheet.Cells["B3"].FormulaR1C1 = "RC[-1]";
            //Set B4 = A5
            worksheet.Cells["B4"].FormulaR1C1 = "R[1]C[-1]";
            //Set B5 = B2 + B3
            worksheet.Cells["B5"].FormulaR1C1 = "R[-3] + R[-2]";

            excelApp.Save("Formula.xlsx");

        }

     
    }
}
