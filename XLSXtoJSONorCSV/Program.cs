using System;
using XLSXtoJSONorCSV.Utility;

namespace XLSXtoJSONorCSV
{
    internal class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("Enter the file name without extension: ");
                string fileName;

                while (true)
                {
                    fileName = Console.ReadLine();
                    if (string.IsNullOrEmpty(fileName))
                        Console.WriteLine("Please enter a valid file name.");
                    else
                        break;
                }

                Console.WriteLine("Enter the desired output format (JSON/XML): ");
                string type = string.Empty;

                while (true)
                {
                    type = Console.ReadLine();
                    if (string.IsNullOrEmpty(type) || (type.ToLower() != "json" && type.ToLower() != "xml"))
                        Console.WriteLine("Please enter a valid output format.");
                    else
                        break;
                }

                CustomUtilityXLS.ConvertXLXStoDesiredFormat(fileName, type);

                Console.WriteLine("Do you want to convert another file? (Y/N)");
                var response = Console.ReadLine();

                if (response.ToLower() == "n")
                    break;

                Console.WriteLine("Press any key to continue...");
                Console.ReadLine();
            }
        }
    }
}
