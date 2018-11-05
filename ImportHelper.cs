using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class ImportHelper
    {

        public static void ImportDataFromArray()
        {
            //Create a data array
            var arrayData = new List<object[]>()
            {
                new object[] {"Name", "BuyIn", "SaleOut", "OrderDate"},
                new object[] {"Orange", 63, 45, "2000-04-25"},
                new object[] {"Apple", 89, 92, "2000-01-12"},
                new object[] {"Banana", 67, 83, "2000-05-21"},
                new object[] {"Pear", 21, 43, "2000-03-09"},
                new object[] {"Grape", 71, 17, "2000-02-03"},
                new object[] {"Watermelon", 33, 55, "2000-03-09"},
            };

            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ImportFromArray");

            //Insert array to worksheet cells
            var dataRange = worksheet.Cells["A1"].LoadFromArrays(arrayData);
            worksheet.Cells["B2:B7"].Style.Numberformat.Format = "#,##0";
            worksheet.Cells["C2:C7"].Style.Numberformat.Format = "#,##0";
            worksheet.Cells["D2:D7"].Style.Numberformat.Format = "yyyy-mm-dd";

            excelApp.Save("ImportFromArray.xlsx");
        }

        public static void ImportDataFromDataTable()
        {
            //Create a data table
            string[] fruits = { "Orange", "Apple", "Banana", "Pear", "Grape", "Watermelon" };
            DataTable dt = new DataTable("Fruit Sales");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Monday", typeof(long));
            dt.Columns.Add("Tuesday", typeof(long));
            dt.Columns.Add("Wednesday", typeof(long));
            dt.Columns.Add("Thursday", typeof(long));
            dt.Columns.Add("Friday", typeof(long));
            dt.Columns.Add("Saturday ", typeof(long));
            dt.Columns.Add("Sunday", typeof(long));

            Random r = new Random();
            for (int row = 0; row < 6; row++)
            {
                var dr = dt.NewRow();
                dr[0] = fruits[row];

                for (int col = 1; col < 8; col++)
                {
                    dr[col] = r.Next(10, 100);
                }

                dt.Rows.Add(dr);
            }

            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ImportFromDataTable");

            //Insert table to worksheet cells
            var dataRange = worksheet.Cells["A1"].LoadFromDataTable(dt, true);
            worksheet.Cells["B2:H7"].Style.Numberformat.Format = "#,##0";

            excelApp.Save("ImportFromDataTable.xlsx");
        }

        public static void ImportDataFromDataReader()
        {
            var excelApp = new Application();

            //Create new worksheet by name
            var worksheet = excelApp.Workbook.Worksheets.Add("ImportFromDataReader");
            
            //Load data from database
            string connectionStr = "Data Source=(local);Initial Catalog=Test;User ID=sa;Password=123456";

            using (SqlConnection sqlConn = new SqlConnection(connectionStr))
            {
                sqlConn.Open();
                using (SqlCommand sqlCmd = new SqlCommand("select column-a, column-b, column-c from table", sqlConn))
                {
                    using (SqlDataReader sqlReader = sqlCmd.ExecuteReader())
                    {
                        //Insert data from sql DataReader to worksheet cells 
                        worksheet.Cells["A1"].LoadFromDataReader(sqlReader, true);                       

                        sqlReader.Close();
                    }
                }

                sqlConn.Close();
            }

            excelApp.Save("ImportFromDataReader.xlsx");
        }

        public static void ImportDataFromIEnumerable()
        {
            //Creat data from list of customized object
            List<Product> fruits = new List<Product>();
            var names = new string[] { "Orange", "Apple", "Banana", "Pear", "Grape", "Watermelon" };
            Random r = new Random();
            for (int i = 0; i < 50; i++)
            {
                fruits.Add(
                    new Product()
                    {
                        Name = names[r.Next(6)],
                        BuyIn = r.Next(50, 200),
                        SaleOut = r.Next(50, 200),
                        OrderDate = new DateTime(2000, 1, 1).AddDays(r.Next(1000)),
                    });
            }

            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ImportFromIEnumerable");

            //Insert IEnumerable<T> to worksheet cells
            var dataRange = worksheet.Cells["A1"].LoadFromCollection(fruits);
            worksheet.Cells["B2:B50"].Style.Numberformat.Format = "#,##0";
            worksheet.Cells["C2:C50"].Style.Numberformat.Format = "#,##0";
            worksheet.Cells["D2:D50"].Style.Numberformat.Format = "yyyy-mm-dd";

            excelApp.Save("ImportFromIEnumerable.xlsx");
        }

        public static void ImportDataFromTextFile()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ImportFromTextFile");

            //Create the format object to describe the text file
            var format = new TextFormat();
            format.Delimiter = '\t'; //Tab
            format.TextQualifier = '"';
            format.SkipLinesBeginning = 1;

            FileInfo textFile = new FileInfo("load-from-text.txt");

            //Insert data from text/csv file to worksheet cells
            var dataRange = worksheet.Cells["A1"].LoadFromText(textFile, format);
            worksheet.Cells["B2:B50"].Style.Numberformat.Format = "#,##0";
            worksheet.Cells["C2:C50"].Style.Numberformat.Format = "#,##0";
            worksheet.Cells["D2:D50"].Style.Numberformat.Format = "yyyy-mm-dd";
            worksheet.Cells["D2:D50"].AutoFitColumns();

            excelApp.Save("ImportFromTextFile.xlsx");
        }

        public class Product
        {
            public string Name { get; set; }
            public int BuyIn { get; set; }
            public int SaleOut { get; set; }
            public DateTime OrderDate { get; set; }           
        }

       
    }
}
