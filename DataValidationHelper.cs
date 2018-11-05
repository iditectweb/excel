using iDiTect.Excel.DataValidation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class DataValidationHelper
    {
        public static void AddIntegerValidation()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ValidationSheet");
                        
            // Add a between value validation 
            var validation = worksheet.Cells["A1"].DataValidation.AddIntegerDataValidation();
            validation.Operator = DataValidationOperator.Between;
            validation.Formula.Value = 1;
            validation.Formula2.Value = 10;
            // Show prompt message
            validation.ShowInputMessage = true;
            validation.PromptTitle = "Enter integer value";
            validation.Prompt = "Value should be between 1 and 10";
            // Show error message
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = DataValidationWarningStyle.Stop;
            validation.ErrorTitle = "An invalid value was entered";
            validation.Error = "Value must be between 1 and 10";

            //You can also add other type validation, like greater than...
            //var validation2 = worksheet.Cells["A2"].DataValidation.AddIntegerDataValidation();
            //validation2.Operator = DataValidationOperator.GreaterThan;
            //validation2.Formula.Value = 9;
            //validation2.ShowInputMessage = true;
            //validation2.PromptTitle = "Enter integer value";
            //validation2.Prompt = "Value should be greater than 9";
            //validation2.ShowErrorMessage = true;
            //validation2.ErrorStyle = DataValidationWarningStyle.Stop;
            //validation2.ErrorTitle = "An invalid value was entered";
            //validation2.Error = "Value must be greater than 9";

            excelApp.Save("IntegerValidation.xlsx");
        }

        public static void AddListValidation()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ValidationSheet");
            
            // Add a list value validation
            var validation = worksheet.DataValidations.AddListValidation("A1");
            // Show prompt message
            validation.ShowInputMessage = true;
            validation.Prompt = "Select or enter a value from the list";
            // Show error message
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = DataValidationWarningStyle.Stop;
            validation.ErrorTitle = "An invalid value was entered";
            validation.Error = "Select a value from the list";

            // Add a list value directly to formula
            for (var i = 1; i <= 5; i++)
            {
                validation.Formula.Values.Add(i.ToString());
            }

            // or you can add list value from cells
            //worksheet.Cells["B1"].Value = 1;
            //worksheet.Cells["B2"].Value = 2;
            //worksheet.Cells["B3"].Value = 3;
            //worksheet.Cells["B4"].Value = 4;
            //worksheet.Cells["B5"].Value = 5;            
            //validation.Formula.ExcelFormula = "B1:B5";

            excelApp.Save("ListValidation.xlsx");
        }

        public static void AddTextLengthValidation()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("ValidationSheet");

            var minLength = 1;
            var maxLength = 10;
            // Add a text length validation 
            var validation = worksheet.Cells["A1"].DataValidation.AddTextLengthDataValidation();
            validation.Formula.Value = minLength;
            validation.Formula2.Value = maxLength;
            // Show prompt message
            validation.ShowInputMessage = true;
            validation.Prompt = "Input charactor length should be between 1 and 10";
            // Show error message
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = DataValidationWarningStyle.Stop;
            validation.ErrorTitle = "Value length is invalid";
            validation.Error = "Value length must be between 1 and 10";

           
            excelApp.Save("TextLengthValidation.xlsx");
        }


    }
}
