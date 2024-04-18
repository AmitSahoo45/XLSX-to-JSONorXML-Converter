using Aspose.Cells;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace XLSXtoJSONorCSV.Utility.Helper
{
    public class Functions
    {
        #region Converting XLSX to JSON
        public static void ConvertXLSXtoJSON(string projectDir, string fileName, string inputFilePath)
        {
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

            string outputFilePath = Path.Combine(projectDir, "Results", "JSON", fileName.Replace(".xlsx", ".json"));

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
                    outputFilePath = Path.Combine(projectDir, "Results", "JSON", $"{fileName.Replace(".xlsx", "")}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.json");
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
        #endregion




        #region Converting XLSX to XML
        public static void ConvertXLSXtoXML(string projectDir, string fileName)
        {
            string inputFilePath = Path.Combine(projectDir, "Data", fileName);
            string outputFilePath = Path.Combine(projectDir, "Results", "XML", fileName.Replace(".xlsx", ".xml"));

            if (File.Exists(outputFilePath))
            {
                Console.WriteLine("File already exists. Do you want to overwrite it? (Y/N)");
                var response = Console.ReadLine();

                if (response.ToLower() == "y")
                {
                    var workbook = new Workbook(inputFilePath);
                    workbook.Save(outputFilePath);
                }
                else
                {
                    outputFilePath = Path.Combine(projectDir, "Results", "XML", $"{fileName.Replace(".xlsx", "")}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xml");
                    var workbook = new Workbook(inputFilePath);
                    workbook.Save(outputFilePath);
                }
            }
            else
            {
                var workbook = new Workbook(Path.Combine(projectDir, "Data", fileName));

                fileName = fileName.Replace(".xlsx", ".xml");
                workbook.Save(Path.Combine(projectDir, "Results", "XML", fileName));
            }

            Console.WriteLine("File created successfully.");
        }
        #endregion
    }
}
