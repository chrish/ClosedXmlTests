using System;
using System.IO;
using ClosedXML;
using ClosedXML.Excel;

namespace ExcelTests
{
    public class Program
    {
        static void Main(string[] args)
        {
            Program p = new Program();

        }

        public Program() 
        {
            string originalSample = "SampleData/Navneliste_sample.xlsx";
            string copiedFile = originalSample.Replace(".xlsx", "_copied.xlsx");

            if (File.Exists(copiedFile)) {
                File.Delete(copiedFile);
            }

            File.Copy(originalSample, copiedFile);

            // Working on a copy of the sample file, opening a specific sheet to see that it handles it correctly
            XLWorkbook wkbook = new XLWorkbook(copiedFile);
            IXLWorksheet sheet = wkbook.Worksheet("Navn");

            string navnCol = GetNavnCol(sheet.FirstRow());

            if (navnCol == null) {
                Console.WriteLine("Can't find the specified column!");
            }

            bool hasData = true;
            int rowCnt = 2;
            do
            {
                if (sheet.Cell(navnCol + rowCnt).IsEmpty())
                {
                    hasData = false;
                }
                else
                {
                    IXLCell cell = sheet.Cell(navnCol + rowCnt);
                    string cellVal = cell.GetValue<string>();
                    if (cellVal.StartsWith("I")) AddComment(cell, "Name starts with I." + Environment.NewLine);
                    if (cellVal.StartsWith("J")) AddComment(cell, "Name starts with J." + Environment.NewLine);
                    if (cellVal.StartsWith("Janne")) AddComment(cell, "This is Janne. She talks. A lot!" + Environment.NewLine);
                    if (cellVal.Contains("bj", StringComparison.OrdinalIgnoreCase)) AddComment(cell, "This name contains bj!" + Environment.NewLine);

                }

                rowCnt++;

            } while (hasData);

            wkbook.Save();
            
        }

        /// <summary>
        /// Get the column index of the column with a Navn header
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public string GetNavnCol(IXLRow row)
        {
            foreach (var r in row.Cells())
            {
                if (r.GetValue<string>().Equals("Navn")) 
                {
                    return r.WorksheetColumn().ColumnLetter();
                }
            }

            return null;
        }

        protected void AddComment(IXLCell cell, string comment) 
        {
            cell.Comment.AddText(comment);
            cell.Comment.Style.Size.SetAutomaticSize();
            cell.Style.Fill.SetBackgroundColor(XLColor.Pink);   
        }
    }

   
}
