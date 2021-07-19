using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using excel = Microsoft.Office.Interop.Excel;
using System.Data.OracleClient;
using System.Data;
using System.Threading;

namespace FinalApplication
{
    class Program
    {
        OracleConnection _Con = null;
        OracleCommand _cmd = null;
        OracleDataAdapter da;
        static void Main(string[] args)
        {

            ChromeOptions chromeOptions = new ChromeOptions();
            IWebDriver driver = new ChromeDriver(chromeOptions);
            driver.Navigate().GoToUrl("https://www.delhisldc.org/Redirect.aspx?Loc=0708");
            driver.Manage().Window.Maximize();
            var dropdown = driver.FindElement(By.Id("ContentPlaceHolder2_cmbdiscom"));
            var c_dropdown = new SelectElement(dropdown);
            c_dropdown.SelectByValue("NDPL");

            Thread.Sleep(1000);
            string ExpectedPath = @"C:\downloads";
            bool fileExists = false;

            //DirectoryInfo D = new DirectoryInfo(ExpectedPath);
            //foreach (FileInfo file in D.GetFiles())
            //{
                //file.Delete();
            //}
           
            chromeOptions.AddUserProfilePreference("Download.default_directory", ExpectedPath);
            driver.FindElement(By.XPath("//*[@id='btnExport']")).Click();
            
            Thread.Sleep(1000);

            excel.Application xlApp = new excel.Application();
            excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExpectedPath+"\\download.xls");
            excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range x1range = xlWorksheet.UsedRange;

            string website;

            string Sql = "";
            int c = x1range.Columns.Count;
            int r = x1range.Rows.Count;
            Program p = new Program();
            DataAccess da = new DataAccess();

            for (int i = 2; i <= c; i++)
            {
                string BlockQuery = "";
                for (int j = 1; j <= r - 1; j++)
                {
                    website = Convert.ToString(x1range.Cells[i][j].value2);
                    BlockQuery = BlockQuery + ", '" + website + " '";

                }
                Sql = "INSERT INTO TP_DDL VALUES(" + BlockQuery.Substring(1, BlockQuery.Length - 1) + " )";
                Console.WriteLine("Insert query created-" +i);
                da.executedata(Sql);
            }
            driver.Quit();
            driver.Dispose();
        }
    }
}

   