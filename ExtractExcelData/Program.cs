using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using Irony.Parsing;
using System.Threading.Tasks;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using DinkToPdf;
using ExtractExcelData.Models;
using System.Linq;

namespace ExtractExcelData
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("------------------- Start ----------------");

            // Read Excel Data
            Console.WriteLine("Reading records from excel started...");
            string filePath = @"D:\Temp\Akshay\ExtractDataFromExcel\ExtractExcelData\Excel\Input\ParcelData.xlsx";
            var parcelRecords = ReadParcelData(filePath);
            Console.WriteLine("Reading records from excel ends...");


            foreach (var parcelRecord in parcelRecords)
            {
                Console.WriteLine($"------------------- Scrapping start for parcel number - {parcelRecord.ParcelNumber}  ----------------");
                WebScrapper(parcelRecord.ParcelNumber);
                Console.WriteLine($"-------------------  Scrapping end for parcel number - {parcelRecord.ParcelNumber}  ----------------");

                Console.WriteLine("Save page as pdf started...");
                SavePageAsPdf(parcelRecord.ParcelNumber, parcelRecord.DocName);
                Console.WriteLine("Save page as pdf ends...");
            }




            // Save Output data to excel
            Console.WriteLine("Getting Payment Histories started...");
            var paymentHistories = GetAllPaymentHistories();
            Console.WriteLine("Getting Payment Histories ends...");


            Console.WriteLine("------------------- Saving Output data to excel stared ----------------");
            SaveToExcel(paymentHistories);
            Console.WriteLine("------------------- Saving Output data to excel ends ----------------");

            Console.WriteLine("------------------- End ----------------");

            Console.ReadLine();
        }

        //---------------------------------------------------------------


        public static void WebScrapper(string parcelNumber)
        {

            // Initialize ChromeDriver
            IWebDriver driver = new ChromeDriver();

            try
            {
                // Navigate to the page
                driver.Navigate().GoToUrl("https://trweb.co.clark.nv.us");

                // Enter Parcel ID
                var parcelInput = driver.FindElement(By.XPath(@"/html/body/div[1]/center/table/tbody/tr[2]/td[2]/table[3]/tbody/tr[2]/td/table/tbody/tr[1]/td[1]/form/table/tbody/tr[1]/td[1]/table/tbody/tr/td[2]/input"));
                parcelInput.SendKeys(parcelNumber);

                // Submit the form
                var submitButton = driver.FindElement(By.Name("Submit"));
                submitButton.Click();

                // Wait for the page to load (adjust time as needed)
                System.Threading.Thread.Sleep(3000);


                // Extract the required data
                var lastPaymentAmount = driver.FindElement(By.XPath("//td[contains(text(), 'Last Payment Amount')]/following-sibling::td")).Text.Trim();
                var lastPaymentDate = driver.FindElement(By.XPath("//td[contains(text(), 'Last Payment Date')]/following-sibling::td")).Text.Trim();
                var fiscalTaxYearPayments = driver.FindElement(By.XPath("//td[contains(text(), 'Fiscal Tax Year Payments')]/following-sibling::td")).Text.Trim();
                var priorCalendarYearPayments = driver.FindElement(By.XPath("//td[contains(text(), 'Prior Calendar Year Payments')]/following-sibling::td")).Text.Trim();
                var currentCalendarYearPayments = driver.FindElement(By.XPath("//td[contains(text(), 'Current Calendar Year Payments')]/following-sibling::td")).Text.Trim();


                // Save data to SQL Server
                SaveDataToSql(parcelNumber, lastPaymentAmount, lastPaymentDate, fiscalTaxYearPayments, priorCalendarYearPayments, currentCalendarYearPayments);

                Console.WriteLine("------------- Scrapping data saved to database -------------");

                // Get the page URL after navigation
                //return driver.Url;

            }
            finally
            {
                // Quit the browser
                driver.Quit();
            }

        }


        static void SaveDataToSql(string parcelNumber, string lastPaymentAmount, string lastPaymentDate, string fiscalTaxYearPayments, string priorCalendarYearPayments, string currentCalendarYearPayments)
        {
            string connectionString = "server=LAPTOP-S2EFS1EF\\SQLEXPRESS;database=ParcelDB;Integrated Security=true;";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "INSERT INTO PaymentHistory1 (ParcelID,LastPaymentAmount, LastPaymentDate, FiscalTaxYearPayments, PriorCalendarYearPayments, CurrentCalendarYearPayments) " +
                               "VALUES (@ParcelID,@LastPaymentAmount, @LastPaymentDate, @FiscalTaxYearPayments, @PriorCalendarYearPayments, @CurrentCalendarYearPayments)";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ParcelID", parcelNumber);
                    command.Parameters.AddWithValue("@LastPaymentAmount", lastPaymentAmount);
                    command.Parameters.AddWithValue("@LastPaymentDate", lastPaymentDate);
                    command.Parameters.AddWithValue("@FiscalTaxYearPayments", fiscalTaxYearPayments);
                    command.Parameters.AddWithValue("@PriorCalendarYearPayments", priorCalendarYearPayments);
                    command.Parameters.AddWithValue("@CurrentCalendarYearPayments", currentCalendarYearPayments);

                    command.ExecuteNonQuery();
                }
            }
        }



        public static List<PaymentHistory> GetAllPaymentHistories()
        {
            string connectionString = "server=LAPTOP-S2EFS1EF\\SQLEXPRESS;database=ParcelDB;Integrated Security=true;";
            var paymentHistories = new List<PaymentHistory>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT * FROM PaymentHistory1";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var paymentHistory = new PaymentHistory
                            {
                                ParcelID = reader["ParcelID"].ToString(),
                                LastPaymentAmount = reader["LastPaymentAmount"].ToString(),
                                LastPaymentDate = reader["LastPaymentDate"].ToString(),
                                FiscalTaxYearPayments = reader["FiscalTaxYearPayments"].ToString(),
                                PriorCalendarYearPayments = reader["PriorCalendarYearPayments"].ToString(),
                                CurrentCalendarYearPayments = reader["CurrentCalendarYearPayments"].ToString()
                            };

                            paymentHistories.Add(paymentHistory);
                        }
                    }
                }
            }

            return paymentHistories;
        }

        static void SaveToExcel(List<PaymentHistory> payments)
        {
            // Specify the directory where you want to save the Excel file
            string directoryPath = @"D:\Temp\Akshay\ExtractDataFromExcel\ExtractExcelData\Excel\Output\";

            // Ensure the directory exists, if not, create it
            if (!System.IO.Directory.Exists(directoryPath))
            {
                System.IO.Directory.CreateDirectory(directoryPath);
            }

            // Full path to the Excel file
            string filePath = System.IO.Path.Combine(directoryPath, "PaymentHistory.xlsx");

            // Create a new Excel workbook
            using (var workbook = new XLWorkbook())
            {
                // Add a worksheet to the workbook
                var worksheet = workbook.Worksheets.Add("PaymentHistory");

                // Add headers to the worksheet
                worksheet.Cell(1, 1).Value = "Parcel ID";
                worksheet.Cell(1, 2).Value = "Last Payment Amount";
                worksheet.Cell(1, 3).Value = "Last Payment Date";
                worksheet.Cell(1, 4).Value = "Fiscal Tax Year Payments";
                worksheet.Cell(1, 5).Value = "Prior Calendar Year Payments";
                worksheet.Cell(1, 6).Value = "Current Calendar Year Payments";

                // Add data to the worksheet
                for (int i = 0; i < payments.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = payments[i].ParcelID;
                    worksheet.Cell(i + 2, 2).Value = payments[i].LastPaymentAmount;
                    worksheet.Cell(i + 2, 3).Value = payments[i].LastPaymentDate; ;
                    worksheet.Cell(i + 2, 4).Value = payments[i].FiscalTaxYearPayments;
                    worksheet.Cell(i + 2, 5).Value = payments[i].PriorCalendarYearPayments;
                    worksheet.Cell(i + 2, 6).Value = payments[i].CurrentCalendarYearPayments;
                }

                // Save the workbook to the specified file path
                workbook.SaveAs(filePath);
            }

            Console.WriteLine($"Data has been successfully saved to {filePath}");
        }
        //-------------------------------------------




        // Extract data from Excel

        public static List<ParcelRecord> ReadParcelData(string filePath)
        {
            var parcelRecords = new List<ParcelRecord>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

                foreach (var row in rows)
                {
                    var parcelRecord = new ParcelRecord
                    {
                        ParcelNumber = row.Cell(1).GetString(),  // A column
                        ParNum = row.Cell(2).GetString(),        // B column
                        DocName = row.Cell(3).GetString(),       // C column
                        DocName1 = row.Cell(4).GetString(),      // D column
                        Stat = row.Cell(5).GetString(),          // E column
                        Rem = row.Cell(6).GetString()            // F column
                    };

                    parcelRecords.Add(parcelRecord);
                }
            }

            return parcelRecords;
        }




        public static async Task SavePageAsPdf(string parcelNumber, string docName)
        {

            string url = $"https://trweb.co.clark.nv.us/print_wep2.asp?Parcel={parcelNumber}";

            // Define the path where PDFs should be saved
            string pdfDirectory = @"D:\Temp\Akshay\ExtractDataFromExcel\ExtractExcelData\pdfs";

            // Ensure the directory exists
            if (!Directory.Exists(pdfDirectory))
            {
                Directory.CreateDirectory(pdfDirectory);
            }

            // Generate a file name based on the parcel number
            string pdfFilePath = Path.Combine(pdfDirectory, $"{docName}.pdf");

            // Set the path to your Chrome or Chromium installation
            var launchOptions = new LaunchOptions
            {
                Headless = true,
                ExecutablePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe", // Update this path accordingly
                DefaultViewport = new ViewPortOptions
                {
                    Width = 1920,
                    Height = 1080
                }
            };

            var browser = await Puppeteer.LaunchAsync(launchOptions);
            var page = await browser.NewPageAsync();

            try
            {
                // Increase timeout duration
                page.DefaultNavigationTimeout = 60000; // Set the timeout to 60 seconds

                // Navigate to the page with increased timeout
                await page.GoToAsync(url, new NavigationOptions
                {
                    Timeout = 60000, // Set the navigation timeout to 60 seconds
                    WaitUntil = new[] { WaitUntilNavigation.Networkidle2 } // Wait until network activity is idle
                });

                // Save the page as a PDF
                await page.PdfAsync(pdfFilePath, new PdfOptions
                {
                    Format = PaperFormat.A4,
                    PrintBackground = true,
                });

                Console.WriteLine($"PDF saved successfully at {pdfFilePath}!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Dispose of resources to ensure everything is cleaned up
                if (page != null)
                {
                    await page.CloseAsync();
                }

                if (browser != null)
                {
                    await browser.CloseAsync();
                }
            }
        }
    }
}
