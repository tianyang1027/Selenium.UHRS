
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System;
using System.IO;
using System.Threading;

namespace Selenium.UHRS
{
    internal class Program
    {
        static void Main()
        {
            string url = "https://www.uhrs.ai/Manage/HitApp/QuickStats?appId=22678&hitAppId=51829&project=QR&easyGrid=judge&startdate=2023-06-01&enddate=2023-06-30&vendorId=-1";
            string filePath = @"C:\Users\v-yangtian\Downloads\output.xlsx";
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");
            IRow row0 = sheet.CreateRow(0);
            IRow row1 = sheet.CreateRow(1);

            using (IWebDriver driver = new EdgeDriver())
            {
                driver.Navigate().GoToUrl(url);
                IWebElement loginBtn = driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[1]/div[2]/div/main/button"));
                loginBtn.Click();
                Thread.Sleep(3000);

                IWebElement userinfoBtn = driver.FindElement(By.XPath("//*[@id=\"tilesHolder\"]/div[1]/div/div/div/div[2]"));
                userinfoBtn.Click();
                Thread.Sleep(10000);

                var dynamicElements = driver.FindElements(By.XPath("//*[@id='easyGridsummary']/div[*]/div[2]"));

                for (int i = 0; i < dynamicElements.Count; i++)
                {
                    var item = dynamicElements[i];
                    var text = item?.Text;
                    row0.CreateCell(i).SetCellValue(GetExcelHeaders()[i]);
                    row1.CreateCell(i).SetCellValue(text);

                }

                using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fileStream);
                }
                Console.WriteLine("output file success!");
                Console.ReadKey();
            }
        }

        private static string[] GetExcelHeaders() => new string[]
        {
            "NumJudges",
            "Total Earnings",
            "Judgment Total",
            "Judging Hours",
            "Judgments / Hour",
            "RTA Accuracy",
            "Spam Accuracy",
            "RTA Total Ratio",
            "Spam Total Ratio",
            "RTA Correct",
            "Spam Correct",
            "Spam Pending"
        };
    }
}
