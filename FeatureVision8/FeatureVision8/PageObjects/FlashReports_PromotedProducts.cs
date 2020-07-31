using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using NUnit.Framework;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Configuration;
using System.Data;
using AventStack.ExtentReports;
using OpenQA.Selenium.Interactions;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Globalization;

namespace FeatureVision8
{
    public class FlashReports_PromotedProducts
    {
        #region Private Variables

        private IWebDriver flashReports_PromotedProducts;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public FlashReports_PromotedProducts(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.flashReports_PromotedProducts = driver;
            test = testReturn;
            homePage = new Home(driver, test);
        }

        public IWebDriver driver
        {
            get { return this.flashReports_PromotedProducts; }
            set { this.flashReports_PromotedProducts = value; }
        }

        ///<summary>
        ///Verify And Select Sub Menu Options of Flash Report
        ///</summary>
        ///<param name="subMenu">Submenu item to be selected</param>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyAndSelectSubMenuOptionsOfFlashReport(string subMenu = "", string clientName = "Procter & Gamble")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//h1[text()='FlashReports']"), "'FlashReports' Screen not present.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='title cursorpointer']/a[text()='Ads And SOV']"), "'Ads And SOV' link not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-media']/i[contains(@class, 'ads-and-share-of-voice')]"), "'Ads And SOV' icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@ng-repeat][1]//div[@class='report-content']/div[contains(@class, 'submenudescription')]"), "'Ads And SOV' description not present.");

            if (clientName.ToLower().Contains("procter"))
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='title cursorpointer']/a[text()='Promoted Products']"), "'Promoted Products' link not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-media']/i[contains(@class, 'promoted-products')]"), "'Promoted Products' icon not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@ng-repeat][2]//div[@class='report-content']/div[contains(@class, 'submenudescription')]"), "'Promoted Products' description not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@ng-repeat][3]//div[@class='report-content']/div[contains(@class, 'submenudescription')]"), "'Retailer Comparison' description not present.");
            }
            else
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@ng-repeat][2]//div[@class='report-content']/div[contains(@class, 'submenudescription')]"), "'Retailer Comparison' description not present.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='title cursorpointer']/a[text()='Retailer Comparison']"), "'Retailer Comparison' link not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-media']/i[contains(@class, 'retailer-comparison')]"), "'Retailer Comparison' icon not present.");

            if (subMenu.ToLower().Equals("ads and sov"))
                driver._click("xpath", "//div[@class='title cursorpointer']/a[text()='Ads And SOV']");
            else if (subMenu.ToLower().Equals("promoted products"))
                driver._click("xpath", "//div[@class='title cursorpointer']/a[text()='Promoted Products']");
            else if (subMenu.ToLower().Equals("retailer comparison"))
                driver._click("xpath", "//div[@class='title cursorpointer']/a[text()='Retailer Comparison']");

            Results.WriteStatus(test, "Pass", "Verified, Sub Menu Options and selected '" + subMenu + "' of Flash Report");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Verify Promoted Products Screen
        ///</summary>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyPromotedProductsScreen()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='flashreport row']//h3"), "'Promoted Products' header not present.");
            Assert.AreEqual("Promoted Products", driver._getText("xpath", "//div[@class='flashreport row']//h3"), "'Promoted Products' header text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'export headerIcon')]"), "Export Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'fa fa-home')]"), "Home Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'all-ads-div')]/span[not(@class)]"), "'Products - Manufacturer' Title not present.");
            Assert.AreEqual("Products - Manufacturer", driver._getText("xpath", "//div[contains(@class, 'all-ads-div')]/span[not(@class)]"), "'Products - Manufacturer' Title text does not match.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvFilterField1002']/div[contains(@class, 'filterHeader')]"), "Pages Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvFilterField1001']/div[contains(@class, 'filterHeader')]"), "'Country' Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[@title='Open Calendar']"), "'Calendar' icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@type='text' and @id='txtCalendar']"), "Calendar Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//li[@heading]"), "Tabs not present.");

            IList<IWebElement> tabsColl = driver._findElements("xpath", "//li[@heading]");
            string[] tabsList = new string[] { "Detail Data", "Promoted Product Images", "Page Images" };

            foreach(string tabName in tabsList)
            {
                bool avail = false;
                foreach(IWebElement tab in tabsColl)
                    if (tab.GetAttribute("heading").ToLower().Equals(tabName.ToLower()))
                    {
                        avail = true;
                        break;
                    }
                Assert.IsTrue(avail, "'" + tabName + "' tab not found.");
            }
            
            Results.WriteStatus(test, "Pass", "Verified, Promoted Products Screen");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Verify Export Options
        ///</summary>
        ///<param name="optionName">Option to be selected</param>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyExportOptions(string optionName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'export headerIcon')]"), "Export Icon not present.");

            if (optionName.ToLower().Contains("excel"))
            {
                driver._click("xpath", "//div[contains(@class, 'export headerIcon')]");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[@id='ulOptionData']/li[text()='Export as Excel']"), "Export DDL option 'Export as Excel' not present.");
                driver._clickByJavaScriptExecutor("//ul[@id='ulOptionData']/li[text()='Export as Excel']");
            }
            else
            {
                driver._click("xpath", "//div[contains(@class, 'export headerIcon')]");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[@id='ulOptionData']/li[text()='Export as PPT']"), "Export DDL option 'Export as PPT' not present.");
                driver._clickByJavaScriptExecutor("//ul[@id='ulOptionData']/li[text()='Export as PPT']");
            }

            driver._waitForElementToBeHidden("xpath", "//div[@id='dvContent1']/div", 600);

            Results.WriteStatus(test, "Pass", "Verified, '" + optionName + "' Export Option");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Capture Data From Detail Data Table
        ///</summary>
        ///<returns></returns>
        public string[,] CaptureDataFromDetailDataTable()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[contains(@class, 'tblRawData')]//div[contains(@class,'headertextdisplayname')]"), "Table Column Headers not present.");
            IList<IWebElement> colHeaderColl = driver._findElements("xpath", "//table[contains(@class, 'tblRawData')]//div[contains(@class,'headertextdisplayname')]");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[contains(@class, 'tblRawData')]//tr[@class]/td[@class]"), "Data Cells in Rows not present.");
            IList<IWebElement> rowCollection = driver._findElements("xpath", "//table[contains(@class, 'tblRawData')]//tr[@class]/td[@class]");

            string[,] dataGrid = new string[11, colHeaderColl.Count];

            for (int j = 0; j < colHeaderColl.Count; j++)
                dataGrid[0, j] = colHeaderColl[j].Text;

            int k = 0;
            for(int i = 1;  i < dataGrid.GetLength(0); i++)
            {
                for (int j = 0; j < dataGrid.GetLength(1); j++, k++)
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", rowCollection[k]);
                    dataGrid[i, j] = rowCollection[k].Text;
                    if (dataGrid[i, j].Contains("/"))
                        dataGrid[i, j] = dataGrid[i, j].Replace("/", "-");
                }
            }

            Results.WriteStatus(test, "Pass", "Captured, Data From Detail Data Table.");
            return dataGrid;
        }

        ///<summary>
        ///Verify Data From Tabular Grid In Exported Excel File
        ///</summary>
        ///<param name="dataGrid">Data Captured from tabular grid</param>
        ///<param name="fileName">Name of Excel File</param>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyDataFromTabularGridInExportedExcelFile(string fileName, string[,] dataGrid)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            int rw = 0;
            int cl = 0;
            string FilePath = "";

            string sourceDir = ExtentManager.ResultsDir + "\\";
            string[] fileEntries = Directory.GetFiles(sourceDir);

            foreach (string fileEntry in fileEntries)
            {
                if (fileEntry.Contains(fileName))
                {
                    FilePath = fileEntry;
                    break;
                }
            }

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(FilePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            int num = xlWorkBook.Sheets.Count;
            for (int s = 1; s <= num; s++)
            {
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(s);
                if (xlWorkSheet.Name.Contains("Detail Data"))
                    break;
            }

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            int rCnt = 1;

            bool found = false;
            while (!found)
            {
                if ((range.Cells[rCnt, 1] as Excel.Range).Text.ToLower().Contains("ad date"))
                    found = true;
                else
                    ++rCnt;

                if (rCnt > 15)
                    break;
            }
            Assert.IsTrue(found, "File data is not correct.");

            for (int i = 0; i < 10; i++, rCnt++)
            {
                for (int j = 0, cCnt = 1; j < 10; j++, cCnt++)
                {
                    Console.Write("i = " + i + ", j = " + j + ", rCnt = " + rCnt + ", cCnt = " + cCnt);
                    Console.WriteLine("Excel : " + (range.Cells[rCnt, cCnt] as Excel.Range).Text + " : dataGrid : " + dataGrid[i, j]);
                    string temp = (range.Cells[rCnt, cCnt] as Excel.Range).Text;
                    while (temp.IndexOf("  ") > -1)
                        temp = temp.Replace("  ", " ");
                    if (temp.Contains("/"))
                        temp = temp.Replace("/", "-");
                    Assert.IsTrue(temp.ToLower().Contains(dataGrid[i, j].ToLower()), "Data Incorrect for (" + i + ", " + j + ")");
                }
            }

            Results.WriteStatus(test, "Pass", "Verified, Data From Tabular Grid In Exported Excel File");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Verify Page/Country DDL
        ///</summary>
        ///<param name="optionName">Option to be selected from DDL</param>
        ///<param name="filterName">Filter to be applied</param>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyPage_CountryDDL(string filterName = "Page", string optionName = "Entire Ad (All Pages)")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvFilterField1002']/div[contains(@class, 'filterHeader')]"), "Pages Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvFilterField1001']/div[contains(@class, 'filterHeader')]"), "'Country' Field not present.");

            if (filterName.ToLower().Contains("page"))
            {
                driver._click("xpath", "//div[@id='dvFilterField1002']/div[contains(@class, 'filterHeader')]");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterField1002']//li"), "Page DDL not present.");
                IList<IWebElement> pageDDLColl = driver._findElements("xpath", "//div[@id='dvFilterField1002']//li");

                bool avail = false;
                foreach(IWebElement pageDDL in pageDDLColl)
                    if (pageDDL.Text.ToLower().Contains(optionName.ToLower()))
                    {
                        avail = true;
                        pageDDL.Click();
                        break;
                    }
                Assert.IsTrue(avail, "'" + optionName + "' not found.");
                Thread.Sleep(2000);

                Assert.IsTrue(driver._waitForElement("xpath", "//tr[@class]//td[@class][21]"), "Page Position Column Cells not present.");
                IList<IWebElement> cellCollection = driver._findElements("xpath", "//tr[@class]//td[@class][21]");

                avail = false;
                foreach (IWebElement cell in cellCollection)
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cell);
                    if (!cell.Text.ToLower().Contains("front"))
                    {
                        avail = true; 
                        break;
                    }
                }

                if (optionName.ToLower().Contains("front"))
                {
                    Assert.IsFalse(avail, "'" + optionName + "' Page Filter not applied successfully.");
                    Results.WriteStatus(test, "Pass", "'" + optionName + "' Page Filter applied successfully.");
                }
                else
                {
                    Assert.IsTrue(avail, "'" + optionName + "' Page Filter not applied successfully.");
                    Results.WriteStatus(test, "Pass", "'" + optionName + "' Page Filter applied successfully.");
                }
            }
            else
            {
                driver._click("xpath", "//div[@id='dvFilterField1001']/div[contains(@class, 'filterHeader')]");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterField1001']//li"), "Country DDL not present.");
                IList<IWebElement> countryDDLColl = driver._findElements("xpath", "//div[@id='dvFilterField1001']//li");

                bool avail = false;
                foreach (IWebElement countryDDL in countryDDLColl)
                    if (countryDDL.Text.ToLower().Contains(optionName.ToLower()))
                    {
                        avail = true;
                        countryDDL.Click();
                        break;
                    }
                Assert.IsTrue(avail, "'" + optionName + "' not found.");
                Thread.Sleep(2000);

                Assert.IsTrue(driver._waitForElement("xpath", "//tr[@class]//td[@class][2]"), "Market Column Cells not present.");
                IList<IWebElement> cellCollection = driver._findElements("xpath", "//tr[@class]//td[@class][2]");

                avail = false;
                foreach (IWebElement cell in cellCollection)
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cell);
                    if (cell.Text.ToLower().Contains("seattle"))
                    {
                        avail = true; ;
                        break;
                    }
                }

                if (optionName.ToLower().Contains("canada"))
                {
                    Assert.IsFalse(avail, "'" + optionName + "' Country Filter not applied successfully.");
                    Results.WriteStatus(test, "Pass", "'" + optionName + "' Country Filter applied successfully.");
                }
                else
                {
                    Assert.IsTrue(avail, "'" + optionName + "' Country Filter not applied successfully.");
                    Results.WriteStatus(test, "Pass", "'" + optionName + "' Country Filter applied successfully.");
                }
            }

            Results.WriteStatus(test, "Pass", "Verify '" + filterName + "' DDL");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Verify Calendar Textbox & Icon
        ///</summary>
        ///<param name="date">Date to be applied</param>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyCalendarTextboxAndIcon(string date = "Latest")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[@title='Open Calendar']"), "'Calendar' icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@type='text' and @id='txtCalendar']"), "Calendar Field not present.");

            DateTime dDate = DateTime.Today;
            if (date.ToLower().Equals("latest"))
                dDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek);
            else
            {
                Assert.IsTrue(DateTime.TryParseExact(date, "MM-dd-yyyy",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out dDate), "Couldn't convert '" + date + " into DateTime'");
                dDate = dDate.AddDays(-(int)DateTime.Today.DayOfWeek);
            }

            string year = dDate.Year.ToString();
            string month = dDate.ToString("MMM", CultureInfo.InvariantCulture);
            string day = dDate.Day.ToString();

            driver._click("xpath", "//span[@title='Open Calendar']");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='ui-datepicker-div']"), "'Calendar' not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='ui-datepicker-div']//div[2]/div[@class='js_selectedText']"), "Year Field not present.");
            if(!driver._getText("xpath", "//div[@id='ui-datepicker-div']//div[2]/div[@class='js_selectedText']").Equals(year))
            {
                driver._click("xpath", "//div[@id='ui-datepicker-div']//div[2]/div[@class='js_selectedText']");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='ui-datepicker-div']//div[2]//li"), "Year DDL not present");
                IList<IWebElement> yearDDLColl = driver._findElements("xpath", "//div[@id='ui-datepicker-div']//div[2]//li");

                bool avail = false;
                foreach(IWebElement yearDDL in yearDDLColl)
                    if (yearDDL.Text.Contains(year))
                    {
                        avail = true;
                        yearDDL.Click();
                        break;
                    }
                Assert.IsTrue(avail, "'" + year + "' Year not found.");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getText("xpath", "//div[@id='ui-datepicker-div']//div[2]/div[@class='js_selectedText']").Equals(year), "'" + year + "' Year not selected.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='ui-datepicker-div']//div[1]/div[@class='js_selectedText']"), "Month Field not present.");
            if (!driver._getText("xpath", "//div[@id='ui-datepicker-div']//div[1]/div[@class='js_selectedText']").Equals(month))
            {
                driver._click("xpath", "//div[@id='ui-datepicker-div']//div[1]/div[@class='js_selectedText']");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='ui-datepicker-div']//div[1]//li"), "Year DDL not present");
                IList<IWebElement> monthDDLColl = driver._findElements("xpath", "//div[@id='ui-datepicker-div']//div[1]//li");

                bool avail = false;
                foreach (IWebElement monthDDL in monthDDLColl)
                    if (monthDDL.Text.Contains(month))
                    {
                        avail = true;
                        monthDDL.Click();
                        break;
                    }
                Assert.IsTrue(avail, "'" + month + "' Month not found.");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getText("xpath", "//div[@id='ui-datepicker-div']//div[1]/div[@class='js_selectedText']").Equals(month), "'" + month + "' Year not selected.");
            }

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='ui-datepicker-div']//td[@available]/a"), "No dates are available in this calendar month.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='ui-datepicker-div']//td[@available]/a[text()='" + day + "']"), "'" + day + "' Day not available for selection.");
            driver._click("xpath", "//div[@id='ui-datepicker-div']//td[@available]/a[text()='" + day + "']");
            Thread.Sleep(2000);

            driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

            date = dDate.ToString("MM/dd/yyyy");
            string dateInTable = dDate.ToString("M/d/yyyy");
            if (dateInTable.Contains("-"))
                dateInTable = dateInTable.Replace("-", "/");

            Assert.IsTrue(driver._waitForElement("xpath", "//tr[@class]//td[@class][1]"), "'Ad Date' Column cells not present.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//tr[@class]//td[@class][1]");

            bool avail1 = true; 
            foreach(IWebElement cell in cellCollection)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cell);
                if (!cell.Text.Contains(dateInTable) &&  cell.Text != "")
                {
                    Console.WriteLine(cell.Text + " : " + dateInTable);
                    avail1 = false;
                    break;
                }
            }
            Assert.IsTrue(avail1, "'" + dateInTable + "' Date not applied successfully.");

            Results.WriteStatus(test, "Pass", "Verified, Calendar Textbox & Icon for date '" + dateInTable + "'.");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Verify Tabs On Promoted Products Page
        ///</summary>
        ///<param name="tabName">Tab to be selected</param>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyTabsOnPromotedProductsPage(string tabName = "Detail Data")
        {
            if (tabName.ToLower().Equals("detail data"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//li[@heading='Detail Data']"), "'Detail Data' Tab not present.");
                driver._click("xpath", "//li[@heading='Detail Data']");
                Thread.Sleep(2000);
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

                Assert.IsTrue(driver._waitForElement("xpath", "//table[contains(@class, 'tblRawData')]"), "'Detail Data' Table not present.");
            }
            else if (tabName.ToLower().Equals("promoted product images"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//li[@heading='Promoted Product Images']"), "'Promoted Product Images' Tab not present.");
                driver._click("xpath", "//li[@heading='Promoted Product Images']");
                Thread.Sleep(2000);
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'tblRawData')]/div[@ng-repeat]"), "'Promoted Product Images' not present.");
            }
            else if (tabName.ToLower().Equals("page images"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//li[@heading='Page Images']"), "'Page Images' Tab not present.");
                driver._click("xpath", "//li[@heading='Page Images']");
                Thread.Sleep(2000);
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//div[contains(@class, 'tblRawData')]/div[@ng-repeat]"), "'Page Images' not present.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Tabs On Promoted Products Page");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Verify Sorting Functionality
        ///</summary>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifySortingFunctionality(string sort = "Ascending")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[contains(@class, 'tblRawData')]//tr/td[contains(@class, 'row_header ')]"), "'Column Headers' not present.");
            IList<IWebElement> colHeaderColl = driver._findElements("xpath", "//table[contains(@class, 'tblRawData')]//tr/td[contains(@class, 'row_header ')]");

            int i = 0;
            foreach(IWebElement colHeader in colHeaderColl)
            {
                IList<IWebElement> colHeaderNameColl = colHeader._findElementsWithinElement("xpath", ".//div[contains(@class, 'divheadertextdisplayname')]");
                if (colHeaderNameColl[0].Text.ToLower().Equals("category"))
                    break;
                ++i;
            }

            Actions action = new Actions(driver);
            driver.MouseHoverByJavaScript(colHeaderColl[i]);
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='divcolumnaction']/span[not(contains(@style, 'display'))]"), "Sort Icon not present.");
            driver.MouseHoverUsingElement("xpath", "//div[@class='divcolumnaction']/span[not(contains(@style, 'display'))]");
            Thread.Sleep(1000);
            driver._clickByJavaScriptExecutor("//div[@class='divcolumnaction']/span[not(contains(@style, 'display'))]");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='divColumnActionMenu']//li"), "Sort by DDL not present.");
            IList<IWebElement> sortDDLColl = driver._findElements("xpath", "//div[@id='divColumnActionMenu']//li");
            bool avail = false;
            foreach(IWebElement sortDDL in sortDDLColl)
                if (sortDDL.Text.ToLower().Contains(sort.ToLower()))
                {
                    sortDDL.Click();
                    avail = true;
                    break;
                }
            Assert.IsTrue(avail, "'Sort by " + sort + "' option not found.");

            driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

            colHeaderColl = driver._findElements("xpath", "//table[contains(@class, 'tblRawData')]//tr/td[contains(@class, 'row_header ')]");
            IList<IWebElement> sortedAngleColl = colHeaderColl[i]._findElementsWithinElement("xpath", ".//div[contains(@class, 'fa-angle')]");
            Assert.AreNotEqual(0, sortedAngleColl.Count, "Sorted Angle Icon not present");
            if (sort.ToLower().Contains("ascend"))
                Assert.IsTrue(sortedAngleColl[0].GetAttribute("class").Contains("up"), "Category Column not sorted in Ascending order.");
            else
                Assert.IsTrue(sortedAngleColl[0].GetAttribute("class").Contains("down"), "Category Column not sorted in Descending order.");

            Assert.IsTrue(driver._waitForElement("xpath", "//table[contains(@class, 'tblRawData')]//tr[@class]/td[@class][5]"), "Column Data Cells not present.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//table[contains(@class, 'tblRawData')]//tr[@class]/td[@class][5]");
            string[] cellDataList = new string[cellCollection.Count];

            for (int j = 0; j < cellCollection.Count; j++)
                cellDataList[j] = cellCollection[j].Text;

            string[] cellDataSorted = new string[cellDataList.Length];
            Array.Copy(cellDataList, cellDataSorted, cellDataList.Length);
            Array.Sort(cellDataSorted);

            if (sort.ToLower().Contains("ascend"))
            {
                Assert.IsTrue(cellDataList.SequenceEqual(cellDataSorted), "Cells of Column not sorted successfully in ascending order.");
                Results.WriteStatus(test, "Pass", "Cells of Column are sorted successfully in ascending order.");
            }
            else if (sort.ToLower().Contains("ascend"))
            {
                Array.Reverse(cellDataSorted);
                Assert.IsTrue(cellDataList.SequenceEqual(cellDataSorted), "Cells of Column not sorted successfully in descending order.");
                Results.WriteStatus(test, "Pass", "Cells of Column are sorted successfully in descending order.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Sorting Functionality");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Verify Product Image On Detail Data Tab
        ///</summary>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyProductImageOnDetailDataTab()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[contains(@class, 'tblRawData')]//tr[@class]/td[not(@class)]/div"), "Prodcuct Image column not present.");
            IList<IWebElement> imageIconColl = driver._findElements("xpath", "//table[contains(@class, 'tblRawData')]//tr[@class]/td[not(@class)]/div");

            driver.MouseHoverByJavaScript(imageIconColl[0]);
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//img"), "Product Image not present.");

            Results.WriteStatus(test, "Pass", "Verified, Product Image On Detail Data Tab");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Verify Items Per Page Functionality
        ///</summary>
        ///<param name="noOfItems">No of Items per page to be selected</param>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyItemsPerPageFunctionality(string noOfItems = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//div[contains(@ng-repeat, 'pageSize')]/button/a"), "Items per page buttons not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[contains(@class, 'tblRawData')]//tr[@class]"), "Records not present.");
            IList<IWebElement> recordsColl = driver._findElements("xpath", "//table[contains(@class, 'tblRawData')]//tr[@class]");
            string selectedItemsNum = driver._getText("xpath", "//div[contains(@class, 'active')]//div[contains(@ng-repeat, 'pageSize')]//button[contains(@class, 'active')]/a");
            Assert.LessOrEqual(recordsColl.Count.ToString(), selectedItemsNum, "Selected Items per page do not match the displayed no. of items.");

            if (selectedItemsNum.Equals(noOfItems))
                Results.WriteStatus(test, "Pass", "'" + noOfItems + "' Items per page is already selected.");
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@ng-repeat, 'pageSize')]//button[@btn-radio='" + noOfItems + "']"), "'" + noOfItems + "' Items per page button not present.");
                driver._click("xpath", "//div[contains(@class, 'active')]//div[contains(@ng-repeat, 'pageSize')]//button[@btn-radio='" + noOfItems + "']");
                homePage.VerifyHomePage();
                Thread.Sleep(2000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[contains(@class, 'active')]//div[contains(@ng-repeat, 'pageSize')]//button[@btn-radio='" + noOfItems + "']", "class").Contains("active"), "'" + noOfItems + "' Items per page button was not selected.");
                recordsColl = driver._findElements("xpath", "//table[contains(@class, 'tblRawData')]//tr[@class]");
                Assert.LessOrEqual(recordsColl.Count.ToString(), noOfItems, "Selected Items per page do not match the displayed no. of items.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Items per page functionality by selecting to '" + noOfItems + "' Items per page.");
            return new FlashReports_PromotedProducts(driver, test);
        }

        ///<summary>
        ///Verify Pagination Functionality
        ///</summary>
        ///<param name="page">Page no. to be selected</param>
        ///<returns></returns>
        public FlashReports_PromotedProducts VerifyPaginationFunctionality(string page = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li/a"), "Pagination not present.");
            IList<IWebElement> paginationColl = driver._findElements("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li/a");

            if (page.ToLower().Contains("first"))
            {
                driver._scrollintoViewElement("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li/a[text()='First']");
                driver._click("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li/a[text()='First']");
                Thread.Sleep(1000);
                paginationColl = driver._findElements("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a[text()='1']"), "First Page is not selected.");
                Assert.IsTrue(paginationColl[0].GetAttribute("class").Contains("disabled"), "First Page Button is not disabled.");
                Assert.IsTrue(paginationColl[1].GetAttribute("class").Contains("disabled"), "Previous Page Button is not disabled.");
            }
            else if (page.ToLower().Contains("last"))
            {
                driver._scrollintoViewElement("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li/a[text()='Last']");
                driver._click("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li/a[text()='Last']");
                Thread.Sleep(1000);
                paginationColl = driver._findElements("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li");
                Assert.IsTrue(paginationColl[paginationColl.Count - 3].GetAttribute("class").Contains("active"), "Last Page is not selected.");
                Assert.IsTrue(paginationColl[paginationColl.Count - 2].GetAttribute("class").Contains("disabled"), "Next Page Button is not disabled.");
                Assert.IsTrue(paginationColl[paginationColl.Count - 1].GetAttribute("class").Contains("disabled"), "Last Page Button is not disabled.");
            }
            else if (page.ToLower().Contains("prev"))
            {
                driver._scrollintoViewElement("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li/a[text()='First']");
                string currentPage = driver._getText("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a");
                int atPage = 0;
                Assert.IsTrue(int.TryParse(currentPage, out atPage), "Couldn't convert '" + currentPage + "' to int.");
                driver._click("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li[contains(@class, 'prev')]/a");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a[text()='" + (atPage - 1).ToString() + "']"), "Previous Page was not selected.");
            }
            else if (page.ToLower().Contains("next"))
            {
                driver._scrollintoViewElement("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li/a[text()='First']");
                string currentPage = driver._getText("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a");
                int atPage = 0;
                Assert.IsTrue(int.TryParse(currentPage, out atPage), "Couldn't convert '" + currentPage + "' to int.");
                driver._click("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li[contains(@class, 'next')]/a");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a[text()='" + (atPage + 1).ToString() + "']"), "Next Page was not selected.");
            }
            else
            {
                bool avail = false;
                IList<IWebElement> paginationCollList = driver._findElements("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li");
                while (!avail && !paginationCollList[paginationCollList.Count - 2].GetAttribute("class").Contains("disabled"))
                {
                    for (int i = 2; i < paginationColl.Count - 2; i++)
                        if (paginationColl[i].Text.Equals(page))
                        {
                            avail = true;
                            paginationColl[i].Click();
                            break;
                        }
                    if (avail)
                        break;
                    paginationColl[paginationColl.Count - 2].Click();
                    Thread.Sleep(1000);
                    paginationColl = driver._findElements("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li/a");
                    paginationCollList = driver._findElements("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li");
                }
                Assert.IsTrue(avail, "'" + page + "' not found.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a[text()='" + page + "']"), "'Page' was not selected.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Pagination functionality by navigating to '" + page + "' page.");
            return new FlashReports_PromotedProducts(driver, test);
        }

        #endregion
    }
}
