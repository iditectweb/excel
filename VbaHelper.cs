using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class VbaHelper
    {
        public static void AddVba()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("VbaSheet");

            //Create a vba project             
            excelApp.Workbook.CreateVBAProject();

            //Add vba code
            //Simplely way to bind vba code
            var sb = new StringBuilder();
            sb.AppendLine("Private Sub Workbook_Open()");
            sb.AppendLine("MsgBox \"VBA Test!\"");
            sb.AppendLine("End Sub");
            excelApp.Workbook.CodeModule.Code = sb.ToString();

            excelApp.Save("vba.xlsm");
        }

        public static void AddVba2()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("VbaSheet");                       

            //Create a vba project             
            excelApp.Workbook.CreateVBAProject();
                       
            //More customized way to bind vba code
            //Add a rectangle shape
            worksheet.Drawings.AddShape("VbaRect", ShapeStyle.Rect);

            var sb = new StringBuilder();
            sb.AppendLine("Public Sub AddTextToRect()");
            sb.AppendLine("[VbaSheet].Shapes(\"VbaRect\").TextEffect.Text = \"Vba sample!\"");
            sb.AppendLine("End Sub");
            //Create a new module and set the code
            var module = excelApp.Workbook.VbaProject.Modules.AddModule("VbaTest");
            module.Code = sb.ToString();
            //Call the newly created sub from the workbook open event
            excelApp.Workbook.CodeModule.Code = "Private Sub Workbook_Open()\r\nAddTextToRect\r\nEnd Sub";

            excelApp.Save("vba.xlsm");
        }
    }
}
