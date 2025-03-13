using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;

public class ExcelReader
{
    public static List<Dictionary<string, string>> ReadHighlightedCells(string filePath)
    {
        List<Dictionary<string, string>> highlightedData = new List<Dictionary<string, string>>();

        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(1); // Read first sheet
            var rows = worksheet.RangeUsed().RowsUsed(); // Get all used rows

            // Read headers (first row)
            var headers = new List<string>();
            foreach (var cell in rows.First().Cells())
            {
               // headers.Add(cell.Value.ToString());
                headers.Add(cell.GetFormattedString());
            }

            // Read data rows and check for highlighted cells
            foreach (var row in rows.Skip(1)) // Skip header row(column names)
            {
                var rowData = new Dictionary<string, string>();
                bool isHighlighted = false;
                int columnIndex = 0;

                foreach (var cell in row.Cells())
                {
                    var rownum= row.RowNumber();
                    // Check if the cell has any background color
                    var bgColor = cell.Style.Fill.BackgroundColor;
                   // if (bgColor.Color.ToArgb()!=XLColor.Transparent.Color.ToArgb())
                   // if (!bgColor.ThemeColor.Equals(XLColor.Transparent))
                     if (bgColor.ColorType != XLColorType.Indexed && bgColor.ColorType != XLColorType.Theme)
                    {
                        isHighlighted = true; // Row is highlighted if any cell has a custom color
                    }
                    rowData[headers[columnIndex]] = cell.GetFormattedString();
                    //rowData[headers[columnIndex]] = cell.Value.ToString();
                    columnIndex++;
                
                }

                // Only add rows that have a highlighted cell
                if (isHighlighted)
                {
                    highlightedData.Add(rowData);
                }
            }
        }

        return highlightedData;
    }
}
