using iDiTect.Excel.Table;
using iDiTect.Excel.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class PivotTablesHelper
    {      

        public static void AddPivotTable()
        {
            var list = GetRandomData();

            var excelApp = new Application();

            //Add data table to sheet1
            var worksheetData = excelApp.Workbook.Worksheets.Add("PivotData");
            //Bind data to cells
            var dataRange = worksheetData.Cells["A1"].LoadFromCollection(list, true, TableStyles.Light1);
            worksheetData.Column(4).Style.Numberformat.Format = "mm-dd-yy";
            dataRange.AutoFitColumns();



            //Add a simple pivot table to sheet2
            var worksheetPivot1 = excelApp.Workbook.Worksheets.Add("Pivot1");
            var pivotTable1 = worksheetPivot1.PivotTables.Add(worksheetPivot1.Cells["A1"], dataRange, "Fruit1");

            //Add row field
            pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
            //Add data field
            pivotTable1.DataFields.Add(pivotTable1.Fields[4]);
            //Customize the pivot table style
            pivotTable1.TableStyle = TableStyles.Light1;
              


            //Add a more complex pivot table to sheet3
            var worksheetPivot2 = excelApp.Workbook.Worksheets.Add("Pivot2");
            var pivotTable2 = worksheetPivot2.PivotTables.Add(worksheetPivot2.Cells["A3"], dataRange, "Fruit2");

            //Add first row field
            pivotTable2.RowFields.Add(pivotTable2.Fields["Name"]);
            //Add another row field
            var rowField = pivotTable2.RowFields.Add(pivotTable2.Fields["OrderDate"]);
            //Group this date field by Years and quaters. 
            rowField.AddDateGrouping(DateGroupBy.Years | DateGroupBy.Quarters);           

            //Add a page field
            pivotTable2.PageFields.Add(pivotTable2.Fields["Title"]);

            //Add the data fields 
            pivotTable2.DataFields.Add(pivotTable2.Fields["BuyIn"]);
            pivotTable2.DataFields.Add(pivotTable2.Fields["SaleOut"]);
            var dataField = pivotTable2.DataFields.Add(pivotTable2.Fields["InCome"]);
            dataField.Format = "#,##0";

            //Show data fields in columns
            pivotTable2.DataOnRows = false;

            excelApp.Save("PivotTable.xlsx");
        }

        public class Product
        {
            public string Name { get; set; }
            public int BuyIn { get; set; }
            public int SaleOut { get; set; }
            public DateTime OrderDate { get; set; }
            public int InCome
            {
                get { return SaleOut - BuyIn; }
            }

            public string Title
            {
                get { return InCome >= 0 ? "Surplus" : "Loss"; }
            }
        }

        private static List<Product> GetRandomData()
        {
            List<Product> fruits = new List<Product>();
            var names = new string[] { "Orange", "Apple", "Banana", "Pear", "Grape", "Watermelon" };
            Random r = new Random();
            for (int i = 0; i < 300; i++)
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
            return fruits;
        }
    }
}
