using System;
using System.Collections.Generic;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string excelFilePath = "sampledatafoodsales.xlsx";

        // Create output folder each time new run
        string outputDirectory = $"output_{DateTime.Now:yyyyMMdd_HHmmss}";

        try
        {

            // Step 1: Read highlighted rows
            List<Dictionary<string, string>> highlightedRows = ExcelReader.ReadHighlightedCells(excelFilePath);
            
            // Step 2: Check if highlighted rows are found
            if (highlightedRows.Count == 0)
            {
                Console.WriteLine("No highlighted rows found.");
            }
            else
            {
                Console.WriteLine($" Highlighted Rows Found: {highlightedRows.Count}");
                // Generate XML only if rows are present
                XmlGenerator.GenerateXml(highlightedRows, outputDirectory);
                Console.WriteLine("XML generation complete!");
            }
            
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

