using iDiTect.Excel.DataValidation;
using iDiTect.Excel.DataValidation.Contracts;
using iDiTect.Excel.Drawing;
using iDiTect.Excel.Drawing.Chart;
using iDiTect.Excel.Licensing;
using iDiTect.Excel.Sparkline;
using iDiTect.Excel.Style;
using iDiTect.Excel.Table;
using iDiTect.Excel.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace iDiTect.Excel.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            //Please repace the trial key from trial-license.txt in download package 
            //This license registration line need to be at very beginning of our other code
            LicenseManager.SetKey("CLBUL-S8CWY-VKDS6-S8N8J-7LU35-GGMYX");

            

            //ChartHelper.AddChart();
            //CommentHelper.AddCommentToSheet();
            //ConditionalFormattingHelper.AddConditionalFormatting();
            CreateAndLoadHelper.CreateNewExcelFile();
            //DataValidationHelper.AddIntegerValidation();
            //FormulaHelper.AddFormula();
            //HyperlinkHelper.AddHyperlinkToSheet();
            //ImportHelper.ImportDataFromArray();
            //PictureHelper.InsertPictureToSheet();
            //PivotTablesHelper.AddPivotTable();
            //ProtectAndLockHelper.ProtectExcel();
            //RowColumnHelper.RowAndColumn();
            //ShapeHelper.AddShapeToSheet();
            //SparklineHelper.AddSparkline();
            //TableHelper.AddTable();
            //VbaHelper.AddVba2();
        }
             


      

    }   
    
}
