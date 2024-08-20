using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExtractDataFromExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
        }



        public List<(string ParcelNumber, string DocName)> LoadParcelNumbers(string excelFilePath)
        {
            var parcelData = new List<(string, string)>();

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // assuming the data is in the first worksheet
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string parcelNumber = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    string docName = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    if (!string.IsNullOrEmpty(parcelNumber) && !string.IsNullOrEmpty(docName))
                    {
                        parcelData.Add((parcelNumber, docName));
                    }
                }
            }

            return parcelData;
        }

    }
}
