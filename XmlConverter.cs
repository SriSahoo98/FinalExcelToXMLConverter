using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

public static class XmlGenerator
{
    public static void GenerateXml(List<Dictionary<string, string>> highlightedRows, string outputDirectory)
    {
        // Ensure the output directory exists, if not then create
        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        // Iterate through each highlighted row and generate a corresponding XML file
        int fileIndex = 1;
        foreach (var row in highlightedRows)
        {
            // Extract values with safe fallback
            string orderId = row.ContainsKey("ID") ? row["ID"] : "unknown";
            string orderDate = row.ContainsKey("Date") ? row["Date"] : "unknown";

            // Extract contact information and parse it
            string contact = row.ContainsKey("Contact") ? row["Contact"] : "unknown";
            var (name, address, city, region) = ParseContact(contact);

            // Extract item details
            string product = row.ContainsKey("Product") ? row["Product"] : "unknown";
            string category = row.ContainsKey("Category") ? row["Category"] : "unknown"; 
            string quantity = row.ContainsKey("Qty") ? row["Qty"] : "0";
            string price = row.ContainsKey("UnitPrice") ? row["UnitPrice"] : "0";
            string total = row.ContainsKey("TotalPrice") ? row["TotalPrice"] : "0";

            //  XML structure
            var shiporder = new XElement("shiporder",
                new XElement("orderperson", "unknown"), // (not available in Excel)
                new XElement("shipto",
                    new XElement("name", name),
                    new XElement("address", address),
                    new XElement("city", city),
                    new XElement("region", region)
                ),
                new XElement("item",
                    new XElement("title", category),
                    new XElement("note", product) , 
                    new XElement("quantity", quantity),
                    new XElement("price", price),
                    new XElement("total", total)
                ),
                
                new XAttribute("orderid", orderId),
                new XAttribute("orderdate", orderDate)
            );

            // Save XML to file
            string outputPath = Path.Combine(outputDirectory, $"output_{fileIndex}.xml");
            shiporder.Save(outputPath);
            Console.WriteLine($"XML file generated: {outputPath}");
            fileIndex++;
        }
    }

    // Use the Contact field and extracts name, address, city, and region
    private static (string name, string address, string city, string region) ParseContact(string contact)
    {
        if (string.IsNullOrWhiteSpace(contact) || !contact.Contains(","))
        {
            return ("unknown", "unknown", "unknown", "unknown");
        }
        else
        {
            // if contact: "University of California San Diego, 9500 Gilman Drive, La Jolla, CA 92093, USA"
            string[] parts = contact.Split(',');
            int length = parts.Length;

            //Extract from the end , "unknown" if not available
            string region = length >= 2 ? parts[length - 2].Trim() : "unknown";  // 2nd last (Region)
            string city = length >= 3 ? parts[length - 3].Trim() : "unknown";    // 3rd last (City)
            string address = length >= 4 ? parts[length - 4].Trim() : "unknown"; // 4th last (Address)
    
            //Everything before these is treated as the name, so join
            string name = length > 4 ? string.Join(",", parts.Take(length - 3)).Trim() : "unknown";

            return (name, address, city, region);
                   
        }    
    }
    
    
}
