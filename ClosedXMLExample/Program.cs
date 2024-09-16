using ClosedXML.Excel;
using System;

class Program
{
    static void Main(string[] args)
    {
        // Create a new workbook
        using (var workbook = new XLWorkbook())
        {
            // Add a worksheet
            var worksheet = workbook.Worksheets.Add("Sample Sheet");

            // Set the value of cell A1
            var cellA1 = worksheet.Cell("A1");
            cellA1.Value = "Sales Report";

            // Set font size, color, and bold
            cellA1.Style.Font.FontSize = 16;
            cellA1.Style.Font.Bold = true;
            cellA1.Style.Font.FontColor = XLColor.White;

            // Merge cells A1 to D1 and apply background color
            worksheet.Range("A1:D1").Merge();
            worksheet.Range("A1:D1").Style.Fill.BackgroundColor = XLColor.Blue;
            worksheet.Range("A1:D1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Add headers to the sheet
            worksheet.Cell("A2").Value = "Product";
            worksheet.Cell("B2").Value = "Quantity";
            worksheet.Cell("C2").Value = "Price";
            worksheet.Cell("D2").Value = "Total";

            // Style the headers (bold, center-aligned, background color)
            var headerRange = worksheet.Range("A2:D2");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.Gray;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            // Add some data
            worksheet.Cell("A3").Value = "Apples";
            worksheet.Cell("B3").Value = 100;
            worksheet.Cell("C3").Value = 1.2;
            worksheet.Cell("D3").FormulaA1 = "=B3*C3"; // Calculate total price

            worksheet.Cell("A4").Value = "Bananas";
            worksheet.Cell("B4").Value = 150;
            worksheet.Cell("C4").Value = 0.8;
            worksheet.Cell("D4").FormulaA1 = "=B4*C4"; // Calculate total price

            worksheet.Cell("A5").Value = "Oranges";
            worksheet.Cell("B5").Value = 200;
            worksheet.Cell("C5").Value = 1.5;
            worksheet.Cell("D5").FormulaA1 = "=B5*C5"; // Calculate total price

            // Apply auto-fit for columns
            worksheet.Columns().AdjustToContents();

            // Save the workbook to a file
            workbook.SaveAs("SalesReport.xlsx");

            Console.WriteLine("Excel file 'SalesReport.xlsx' created successfully!");
        }
    }
}
