using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using System.Xml;
using Aspose.Cells;

namespace XLSXtoJSONorCSV.Utility
{
    internal class CustomUtilityXLS
    {
        public static void ConvertXLSXtoDesiredFormat(string fileName, string to)
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

                if (to.ToLower() == "json")
                    Helper.Functions.ConvertXLSXtoJSON(projectDir, fileName, inputFilePath);
                else if (to.ToLower() == "xml")
                    Helper.Functions.ConvertXLSXtoXML(projectDir, fileName);
                else
                    throw new ArgumentException("Invalid output format. Please enter JSON or XML.");
            }
            catch (Exception ex)
            {
                Console.Write("An error occurred while converting the file - ");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
