using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;


namespace XLSXtoJSONorCSV.Utility
{
    internal class CustomUtilityXLS
    {
        public static void ConvertXLSXtoJSON(string projectDir, string fileName, List<Dictionary<string, object>> data)
        {
            string outputFilePath = Path.Combine(projectDir, "Results", fileName.Replace(".xlsx", ".json"));

            if (File.Exists(outputFilePath))
            {
                Console.WriteLine("File already exists. Do you want to overwrite it? (Y/N)");
                var response = Console.ReadLine();

                if (response.ToLower() == "y")
                {
                    File.WriteAllText(outputFilePath, JsonConvert.SerializeObject(data));
                    Console.WriteLine("File overwritten successfully.");
                }
                else
                {
                    outputFilePath = Path.Combine(projectDir, "Results", $"{fileName.Replace(".xlsx", "")}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.json");
                    File.WriteAllText(outputFilePath, JsonConvert.SerializeObject(data));
                    Console.WriteLine("File created successfully.");
                }
            }
            else
            {
                File.WriteAllText(outputFilePath, JsonConvert.SerializeObject(data));
                Console.WriteLine("File created successfully.");
            }
        }

        public static void ConvertXLSXtoXML()
        {
        }

        public static void ConvertXLXStoDesiredFormat(string fileName, string to)
        {
            try
            {
                if (string.IsNullOrEmpty(fileName))
                    throw new ArgumentNullException("fileName");

                string projectDir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                fileName += fileName.EndsWith(".xlsx") ? "" : ".xlsx";

                string inputFilePath = Path.Combine(projectDir, "Data", fileName);

                if (!File.Exists(inputFilePath))
                    throw new FileNotFoundException("File with name " + fileName + " not found.");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var excelPackage = new ExcelPackage(new FileInfo(inputFilePath));
                var worksheet = excelPackage.Workbook.Worksheets[0];

                var rowCount = worksheet.Dimension.Rows;
                var columnsCount = worksheet.Dimension.Columns;

                var headers = new List<string>();
                for (int col = 1; col <= columnsCount; col++)
                    headers.Add(worksheet.Cells[1, col].Value.ToString());

                var data = new List<Dictionary<string, object>>();
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, object>();
                    for (int col = 1; col <= columnsCount; col++)
                        rowData[headers[col - 1]] = worksheet.Cells[row, col].Value;

                    data.Add(rowData);
                }

                if (to.ToLower() == "json")
                    ConvertXLSXtoJSON(projectDir, fileName, data);
                else if (to.ToLower() == "xml")
                    ConvertXLSXtoXML();
                else
                    throw new ArgumentException("Invalid output format. Please enter JSON or CSV.");

                excelPackage.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
