using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace ExtractExcelData
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("------------------- Start ----------------");

            var parcelData = LoadParcelNumbers("D:\\Temp\\Akshay\\ExtractDataFromExcel\\ExtractExcelData\\ExcelData_Aug192024.xlsx");

            InsertParcelDataIntoDatabase(parcelData);



            foreach (var data in parcelData)
            {
                Console.WriteLine(data.ParcelNumber);
                // InsertParcelDataIntoDatabase(data.ParcelNumber, data.ParNumber, data.DocName, data.DocName1);
                //FetchAndSaveParcelData(data.ParcelNumber, data.DocName);

            }

            Console.WriteLine("------------------- End ----------------");

            Console.ReadLine();
        }


        // Extract Parcel Number from Excel
        public static List<Parcel> LoadParcelNumbers(string excelFilePath)
        {
            //var parcelData = new List<(string, string)>();
            var parcelData = new List<Parcel>();

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int rowCount = worksheet.Dimension.Rows;
                Parcel parcel = new Parcel();
                for (int row = 2; row <= rowCount; row++)
                {
                    parcel.ParcelNumber = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    parcel.ParNumber = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    parcel.DocName = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    parcel.DocName1 = worksheet.Cells[row, 4].Value?.ToString().Trim();

                    if (!string.IsNullOrEmpty(parcel.ParcelNumber) && !string.IsNullOrEmpty(parcel.DocName))
                    {
                        parcelData.Add(parcel);
                    }
                }
            }

            return parcelData;
        }


        // Save data in Database

        public static void InsertParcelDataIntoDatabase(List<Parcel> parcelData)
        {
            try
            {


                string connectionString = "server=LAPTOP-S2EFS1EF\\SQLEXPRESS;database=ParcelDB;trusted_connection=true;";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = @"
                INSERT INTO ParcelTable (ParcelNumber, ParNumber, DocName, DocName1) 
                VALUES (@ParcelNumber, @ParNumber, @DocName, @DocName1)";

                        // Add parameters
                        command.Parameters.Add("@ParcelNumber", SqlDbType.VarChar);
                        command.Parameters.Add("@ParNumber", SqlDbType.VarChar);
                        command.Parameters.Add("@DocName", SqlDbType.VarChar);
                        command.Parameters.Add("@DocName1", SqlDbType.VarChar);

                        foreach (var record in parcelData)
                        {
                            command.Parameters["@ParcelNumber"].Value = record.ParcelNumber;
                            command.Parameters["@ParNumber"].Value = record.ParNumber;
                            command.Parameters["@DocName"].Value = record.DocName;
                            command.Parameters["@DocName1"].Value = record.DocName1;

                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
            }
        }
    }
}
