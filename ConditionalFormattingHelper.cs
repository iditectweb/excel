using iDiTect.Excel.ConditionalFormatting;
using iDiTect.Excel.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
   public static class ConditionalFormattingHelper
    {
        public static void AddConditionalFormatting()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ConditionalFormatting");

            // Create 10 columns of samples data
            for (int col = 1; col < 11; col++)
            {              
                worksheet.Cells[1, col].Value = "Sample " + col;

                for (int row = 2; row < 12; row++)
                { 
                    worksheet.Cells[row, col].Value = row - 1;
                }
            }

            ExcelAddress address1 = new ExcelAddress("A2:A11");
            var rule1 = worksheet.ConditionalFormatting.AddTwoColorScale(address1);
            rule1.LowValue.Color = System.Drawing.Color.Red;
            rule1.HighValue.Color = System.Drawing.Color.Purple;
            //similar
            //var rule1 = worksheet.ConditionalFormatting.AddThreeColorScale(address1);
         
            ExcelAddress address2 = new ExcelAddress(2, 2, 11, 2);  //="B2:B10"
            var rule2 = worksheet.ConditionalFormatting.AddAboveAverage(address2);
            rule2.Style.Font.Strike = true;
            //similar
            //var rule2 = worksheet.ConditionalFormatting.AddAboveOrEqualAverage(address2);
            //var rule2 = worksheet.ConditionalFormatting.AddAboveStdDev(address2);
            //var rule2 = worksheet.ConditionalFormatting.AddBelowAverage(address2);
            //var rule2 = worksheet.ConditionalFormatting.AddBelowOrEqualAverage(address2);
            //var rule2 = worksheet.ConditionalFormatting.AddBelowStdDev(address2);

            ExcelAddress address3 = new ExcelAddress("C2:C11");
            var rule3 = worksheet.ConditionalFormatting.AddBottomPercent(address3);
            //Highlight the cells whose values fall in the bottom 30%
            rule3.Rank = 30;
            rule3.Style.Font.Color.Color = System.Drawing.Color.Red;
            //similar
            //var rule3 = worksheet.ConditionalFormatting.AddBottom(address3);
            //var rule3 = worksheet.ConditionalFormatting.AddTop(address3);
            //var rule3 = worksheet.ConditionalFormatting.AddTopPercent(address3);

            ExcelAddress address4 = new ExcelAddress("D2:D11");
            var rule4 = worksheet.ConditionalFormatting.AddBeginsWith(address4);
            rule4.Text = "3";
            rule4.Style.Fill.BackgroundColor.Color = System.Drawing.Color.Red;
            //similar
            //var rule4 = worksheet.ConditionalFormatting.AddEndsWith(address4);
            //var rule4 = worksheet.ConditionalFormatting.AddContainsText(address4);
            //var rule4 = worksheet.ConditionalFormatting.AddNotContainsText(address4);

            ExcelAddress address5 = new ExcelAddress("E2:E11");
            var rule5 = worksheet.ConditionalFormatting.AddBetween(address5);
            rule5.Formula = "4";
            rule5.Formula2 = "6";
            rule5.Style.Font.Color.Color = System.Drawing.Color.Red;
            //similar
            //var rule5 = worksheet.ConditionalFormatting.AddNotBetween(address5);

            ExcelAddress address6 = new ExcelAddress("F2:F11");
            var rule6 = worksheet.ConditionalFormatting.AddGreaterThan(address6);
            rule6.Formula = "5";
            rule6.Style.Font.Color.Color = System.Drawing.Color.Red;
            //similar
            //var rule6 = worksheet.ConditionalFormatting.AddGreaterThanOrEqual(address6);
            //var rule6 = worksheet.ConditionalFormatting.AddLessThan(address6);
            //var rule6 = worksheet.ConditionalFormatting.AddLessThanOrEqual(address6);
            //var rule6 = worksheet.ConditionalFormatting.AddEqual(address6);
            //var rule6 = worksheet.ConditionalFormatting.AddNotEqual(address6);

            ExcelAddress address7 = new ExcelAddress("G2:G11");
            var rule7 = worksheet.ConditionalFormatting.AddContainsBlanks(address7);
            rule7.Style.Font.Color.Color = System.Drawing.Color.Red;
            //similar
            //var rule7 = worksheet.ConditionalFormatting.AddContainsErrors(address7);
            //var rule7 = worksheet.ConditionalFormatting.AddNotContainsBlanks(address7);
            //var rule7 = worksheet.ConditionalFormatting.AddNotContainsErrors(address7);
            //var rule7 = worksheet.ConditionalFormatting.AddDuplicateValues(address7);

            ExcelAddress address8 = new ExcelAddress("H2:H11");
            var rule8 = worksheet.ConditionalFormatting.AddExpression(address8);
            rule8.Formula = "IF(H2-H3<=0,1,0)";
            rule8.Style.Font.Color.Color = System.Drawing.Color.Red;

            ExcelAddress address9 = new ExcelAddress("I2:I11");
            var rule9 = worksheet.ConditionalFormatting.AddThreeIconSet(address9, ConditionalFormatting3IconsSetType.TrafficLights1);
            //similar
            //var rule9 = worksheet.ConditionalFormatting.AddFourIconSet(address9, ConditionalFormatting4IconsSetType.TrafficLights);
            //var rule9 = worksheet.ConditionalFormatting.AddFiveIconSet(address9, ConditionalFormatting5IconsSetType.Arrows);

            ExcelAddress address10 = new ExcelAddress("J2:J11");
            var rule10 = worksheet.ConditionalFormatting.AddDatabar(address10, System.Drawing.Color.Red);


            worksheet.Cells[1, 11].Value = "Sample 11";
            for (int row = 2; row < 12; row++)
            {
                worksheet.Cells[row, 11].Value = DateTime.Now.AddDays(-row+1);
                worksheet.Cells[row, 11].Style.Numberformat.Format = "mm-dd-yy";
            }

            ExcelAddress address11 = new ExcelAddress("K2:K11");
            var rule11 = worksheet.ConditionalFormatting.AddLast7Days(address11);
            rule11.Style.Font.Color.Color = System.Drawing.Color.Red;
            //similar
            //var rule11 = worksheet.ConditionalFormatting.AddLastMonth(address11);
            //var rule11 = worksheet.ConditionalFormatting.AddLastWeek(address11);
            //var rule11 = worksheet.ConditionalFormatting.AddNextMonth(address11);
            //var rule11 = worksheet.ConditionalFormatting.AddNextWeek(address11);
            //var rule11 = worksheet.ConditionalFormatting.AddThisMonth(address11);
            //var rule11 = worksheet.ConditionalFormatting.AddThisWeek(address11);
            //var rule11 = worksheet.ConditionalFormatting.AddToday(address11);
            //var rule11 = worksheet.ConditionalFormatting.AddTomorrow(address11);

            excelApp.Save("conditionalFormatting.xlsx");
        }
       
    }
}
