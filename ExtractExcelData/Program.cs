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

namespace ExtractExcelData
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("------------------- Start ----------------");

            var parcelData = LoadParcelNumbers("D:\\Temp\\Akshay\\ExtractDataFromExcel\\ExtractExcelData\\ExcelData_Aug192024.xlsx");

            InsertParcelDataIntoDatabase(parcelData);




            var options = new ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", Path.GetFullPath(@"D:\Temp\Akshay\ExtractDataFromExcel\ExtractExcelData\temp"));

            using (IWebDriver driver = new ChromeDriver(options))
            {
                foreach (var data in parcelData)
                {
                    Console.WriteLine(data.ParcelNumber);
                    // InsertParcelDataIntoDatabase(data.ParcelNumber, data.ParNumber, data.DocName, data.DocName1);
                    FetchAndSaveParcelData(driver, data.ParcelNumber);

                }

                // Close the browser
                driver.Quit();
            }


            // Retrieve from database
            var paymentData = RetrieveFromDatabase();

            // Save to Excel
            SaveToExcel(paymentData);

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

                for (int row = 2; row <= rowCount; row++)
                {
                    Parcel parcel = new Parcel();
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


        // Fetch And Save Parcel Data

        public static void FetchAndSaveParcelData(IWebDriver driver, string parcelNumber)
        {

            // Set a longer command timeout if needed
            var options = new ChromeOptions();
            options.ScriptTimeout = TimeSpan.FromMinutes(3);

            // Set timeouts
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromMinutes(3); // Increase the timeout for page load
            driver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromSeconds(60);

            // Step 1: Open the URL
            driver.Navigate().GoToUrl("https://treasurer.pinal.gov/ParcelInquiry/");

            // Step 2: Wait for a specific element to be visible
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMinutes(3));
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("k-input-inner")));

            // Step 3: Enter Parcel Number and Submit
            var parcelInput = driver.FindElement(By.ClassName("k-input-inner"));
            parcelInput.SendKeys(parcelNumber);
            var submitButton = driver.FindElement(By.XPath(@"/html/body/div[1]/section/table/tbody/tr/td[2]/div/div[1]/form/div/input[1]"));
            submitButton.Click();

            // Step 4: Wait for the payment history link to be clickable
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(@"/html/body/div[1]/section/table/tbody/tr/td[1]/ul/li[1]/ul/li[4]/a")));

            var paymentHistory = driver.FindElement(By.XPath(@"/html/body/div[1]/section/table/tbody/tr/td[1]/ul/li[1]/ul/li[4]/a"));
            paymentHistory.Click();

            // Step 5: Year Dropdown
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(@"/html/body/div[1]/section/table/tbody/tr/td[2]/form/table/tbody/tr/td[2]/span/span[2]")));
            var yearDropdown = driver.FindElement(By.XPath(@"/html/body/div[1]/section/table/tbody/tr/td[2]/form/table/tbody/tr/td[2]/span/span[2]"));
            yearDropdown.Click();

            // Step 6: Select Year Option using JavaScript
            var yearOption = driver.FindElement(By.XPath(@"/html/body/div[2]/div/div/div[2]/ul/li[3]/span"));
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            jsExecutor.ExecuteScript("arguments[0].click();", yearOption);

            // After the page loads, save it as a PDF
            //SavePageAsPdf(driver.Url, parcelNumber).Wait();

            // Step 7: Extract Data
            wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("k-master-row")));
            var rows = driver.FindElements(By.ClassName("k-master-row"));

            foreach (var row in rows)
            {
                var batchNumber = row.FindElement(By.CssSelector("td:nth-child(3)")).Text;
                var paymentDate = row.FindElement(By.CssSelector("td:nth-child(4)")).Text;
                var interestDate = row.FindElement(By.CssSelector("td:nth-child(5)")).Text;
                var payee = row.FindElement(By.CssSelector("td:nth-child(6)")).Text;
                var batchAmount = row.FindElement(By.CssSelector("td:nth-child(7)")).Text;

                // Save Data to Database 
                SaveToDatabase(new PaymentData
                {
                    ParcelNumber = parcelNumber,
                    BatchNumber = batchNumber,
                    PaymentDate = paymentDate,
                    InterestDate = interestDate,
                    Payee = payee,
                    BatchAmount = batchAmount
                });
            }
        }


        public static async Task SavePageAsPdf(string url, string parcelNumber)
        {
            // Define the path where PDFs should be saved
            string pdfDirectory = @"D:\Temp\Akshay\ExtractDataFromExcel\ExtractExcelData\pdfs";

            // Ensure the directory exists
            if (!Directory.Exists(pdfDirectory))
            {
                Directory.CreateDirectory(pdfDirectory);
            }

            // Generate a file name based on the parcel number
            string pdfFilePath = Path.Combine(pdfDirectory, $"{parcelNumber}_payment_history.pdf");

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



        // Save data to excel
        static void SaveToExcel(List<PaymentData> payments)
        {
            // Specify the directory where you want to save the Excel file
            string directoryPath = @"D:\Temp\Akshay\ExtractDataFromExcel\ExtractExcelData\Excel";

            // Ensure the directory exists, if not, create it
            if (!System.IO.Directory.Exists(directoryPath))
            {
                System.IO.Directory.CreateDirectory(directoryPath);
            }

            // Full path to the Excel file
            string filePath = System.IO.Path.Combine(directoryPath, "Payments.xlsx");

            // Create a new Excel workbook
            using (var workbook = new XLWorkbook())
            {
                // Add a worksheet to the workbook
                var worksheet = workbook.Worksheets.Add("Payments");

                // Add headers to the worksheet

                worksheet.Cell(1, 1).Value = "Parcel Number";
                worksheet.Cell(1, 2).Value = "Batch #";
                worksheet.Cell(1, 3).Value = "Payment Date";
                worksheet.Cell(1, 4).Value = "Interest Date";
                //worksheet.Cell(1, 4).Value = "Payee";
                worksheet.Cell(1, 5).Value = "Batch Amount";

                // Add data to the worksheet
                for (int i = 0; i < payments.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = payments[i].ParcelNumber;
                    worksheet.Cell(i + 2, 2).Value = payments[i].BatchNumber;
                    worksheet.Cell(i + 2, 3).Value = payments[i].PaymentDate;
                    worksheet.Cell(i + 2, 4).Value = payments[i].InterestDate;
                    //worksheet.Cell(i + 2, 4).Value = payments[i].Payee;
                    worksheet.Cell(i + 2, 5).Value = payments[i].BatchAmount;
                }

                // Save the workbook to the specified file path
                workbook.SaveAs(filePath);
            }

            Console.WriteLine($"Data has been successfully saved to {filePath}");
        }


        // Extract rows from web link
        public static List<PaymentData> ExtractData(IWebDriver driver, string parcelNumber)
        {
            // Find the table rows
            //var rows = driver.FindElements(By.CssSelector("tbody tr"));//k-master-row

            var rows = driver.FindElements(By.ClassName("k-master-row"));

            // Create a list to hold the extracted data
            List<PaymentData> payments = new List<PaymentData>();

            foreach (var row in rows)
            {
                // Extract the relevant data from each cell
                var batchNumber = row.FindElement(By.CssSelector("td:nth-child(3)")).Text;
                var paymentDate = row.FindElement(By.CssSelector("td:nth-child(4)")).Text;
                var interestDate = row.FindElement(By.CssSelector("td:nth-child(5)")).Text;
                var payee = row.FindElement(By.CssSelector("td:nth-child(6)")).Text;
                var batchAmount = row.FindElement(By.CssSelector("td:nth-child(7)")).Text;

                // Add the extracted data to the list
                payments.Add(new PaymentData
                {
                    ParcelNumber = parcelNumber,
                    BatchNumber = batchNumber,
                    PaymentDate = paymentDate,
                    InterestDate = interestDate,
                    Payee = payee,
                    BatchAmount = batchAmount
                });
            }


            return payments;
        }


        // Save Data to Database 
        public static void SaveToDatabase(PaymentData payment)
        {
            // Connection string to your SQL Server database
            string connectionString = "Data Source=LAPTOP-S2EFS1EF\\SQLEXPRESS;Initial Catalog=ParcelDB;Integrated Security=true";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();


                // SQL query to insert data
                string query = "INSERT INTO PaymentHistory (ParcelNumber, BatchNumber, PaymentDate, InterestDate, BatchAmount) " +
                               "VALUES (@ParcelNumber, @BatchNumber, @PaymentDate, @InterestDate, @BatchAmount)";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // Add parameters to the command
                    command.Parameters.AddWithValue("@ParcelNumber", payment.ParcelNumber);
                    command.Parameters.AddWithValue("@BatchNumber", payment.BatchNumber);
                    command.Parameters.AddWithValue("@PaymentDate", payment.PaymentDate);
                    command.Parameters.AddWithValue("@InterestDate", payment.InterestDate);
                    command.Parameters.AddWithValue("@BatchAmount", payment.BatchAmount);

                    // Execute the query
                    command.ExecuteNonQuery();
                }


                Console.WriteLine("Data has been successfully saved to the database.");
            }
        }


        // Retrieve Data From Database
        public static List<PaymentData> RetrieveFromDatabase()
        {
            // Connection string to your SQL Server database
            string connectionString = "Data Source=LAPTOP-S2EFS1EF\\SQLEXPRESS;Initial Catalog=ParcelDB;Integrated Security=true";

            // List to hold the retrieved data
            List<PaymentData> payments = new List<PaymentData>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // SQL query to retrieve data
                string query = "SELECT ParcelNumber, BatchNumber, PaymentDate, InterestDate, BatchAmount FROM PaymentHistory";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            // Retrieve each column's value
                            string parcelNumber = reader["ParcelNumber"].ToString();
                            string batchNumber = reader["BatchNumber"].ToString();
                            string paymentDate = reader["PaymentDate"].ToString();
                            string interestDate = reader["InterestDate"].ToString();
                            string batchAmount = reader["BatchAmount"].ToString();

                            // Add the retrieved data to the list
                            payments.Add(new PaymentData
                            {
                                ParcelNumber = parcelNumber,
                                BatchNumber = batchNumber,
                                PaymentDate = paymentDate,
                                InterestDate = interestDate,
                                BatchAmount = batchAmount
                            });
                        }
                    }
                }
            }

            return payments;
        }
    }




}
