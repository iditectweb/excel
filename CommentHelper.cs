using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iDiTect.Excel.Demo
{
    public static class CommentHelper
    {
        public static void AddCommentToSheet()
        {
            var excelApp = new Application();

            var worksheet = excelApp.Workbook.Worksheets.Add("CommentSheet");

            //Add simple comment
            var comment = worksheet.Cells["A1"].AddComment("Comment1:", "Author");
            comment.Text = "Single line comment.";
            comment.Font.Bold = true;

            //Add multiple lines comment
            comment = worksheet.Cells["A2"].AddComment("Comment2:", "Author");
            var rt = comment.RichText.Add("\r\nFrist line in comment.\r\nSencond line.");
            rt.Bold = false;

            //Add auto fit width comment
            comment = worksheet.Cells["A3"].AddComment("Comment3:", "Author");
            rt = comment.RichText.Add("\r\nFrist line in comment, this is a long line.\r\nSencond line.");
            comment.AutoFit = true;


            //Add a comment using the Comment collection
            comment = worksheet.Comments.Add(worksheet.Cells["A4"], "Comment4", "Author");
            //Set the size and position. The position is available only when the comment is visible.          
            comment.From.Row = 5;
            comment.To.Row = 10;
            comment.From.Column = 3;
            comment.To.Column = 10;  
            //Set back color         
            comment.BackgroundColor = Color.White;
            //Change comment font by font name            
            rt = comment.RichText.Add("\r\nThis");
            rt.FontName = "Times New Roman";
            //Set font color
            rt.Color = Color.Red;
            rt = comment.RichText.Add(" is");
            rt.Color = Color.FromArgb(0, 0, 128, 0);
            rt = comment.RichText.Add(" colorful");
            rt.Color = Color.Blue;
            rt = comment.RichText.Add(" line.");
            rt.Color = Color.Black;
            

            excelApp.Save("comment.xlsx");
        }
    }
}
