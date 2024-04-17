using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;


namespace XLSXtoJSONorCSV.Utility
{
    internal class CustomUtilityXLS
    {
        public static void ConvertXLSXtoJSON()
        {

        }

        public static void ConvertXLSXtoCSV()
        {

        }

        public static void ConvertXLXStoDesiredFormat(string fileName)
        {
            try
            {
                if (string.IsNullOrEmpty(fileName))
                    throw new ArgumentNullException("fileName");

                string projectDir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
                fileName += fileName.EndsWith(".xlsx") ? "" : ".xlsx";

                string inputFilePath = Path.Combine(projectDir, "Data", fileName);

                if (!File.Exists(inputFilePath))
                    throw new FileNotFoundException("File not found", inputFilePath);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var excelPackage = new ExcelPackage(new FileInfo(inputFilePath));
                var worksheet = excelPackage.Workbook.Worksheets[0];

                var rowCount = worksheet.Dimension.Rows;
                var columnsCount = worksheet.Dimension.Columns;

                var data = new object[rowCount, columnsCount];

                for (int row = 1; row <= rowCount; row++)
                    for (int col = 1; col <= columnsCount; col++)
                        data[row - 1, col - 1] = worksheet.Cells[row, col].Value;

                // check if the there is a file in the Results folder with the same name as the value stored in fileName variable
                // if it exists, ask the user if they want to overwrite the file
                // if yes, overwrite the file
                // if no, create a new file with a naming convention fileName + DateTime.Now.ToString("yyyyMMddHHmmss")
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

                excelPackage.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
