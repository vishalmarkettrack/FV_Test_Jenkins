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

namespace FeatureVision8
{
    public class Calendar
    {
        #region Private Variables

        private IWebDriver calendar;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public Calendar(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.calendar = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.calendar; }
            set { this.calendar = value; }
        }

        ///<summary>
        ///Verify Calendar Screen
        ///</summary>
        ///<returns></returns>
        public Calendar VerifyCalendarScreen(string clientName = "Procter & Gamble", string searchName = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//h1[text()='Calendar']"), "Calendar Screen not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='title cursorpointer']/a"), "Custom Calendar link not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-media']/i"), "Custom Calendar icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-content']/div[contains(@class, 'submenudescription')]"), "Retailer Circular Strategy description not present.");

            driver._click("xpath", "//div[@class='title cursorpointer']/a");
            Thread.Sleep(1000);

            driver._waitForElementToBeHidden("xpath", "//div[@class='ProcessLoader'], 45");

            Assert.IsTrue(driver._waitForElement("xpath", "//img[@src='/Images/isc-NumeratorLogo.png']", 15), "Numerator Logo not found on Home Page.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='report-content']/h1"), "Madlib Search header text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[contains(@class, 'filler')]"), "Madlib Prompt Summary text not present.");

            if (searchName != "")
                Assert.IsTrue(driver._getText("xpath", "//div[@class='report-content']/h1").ToLower().Contains(searchName.ToLower()), "Search Name '" + searchName + "' does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='1']"), "Madlib Search Parameter 'Any Product' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='2']"), "Madlib Search Parameter 'Any Account' not present.");
            if(!clientName.ToLower().Contains("australia"))
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='3']"), "Madlib Search Parameter 'Any Market' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='4']"), "Madlib Search Parameter 'Any Date' not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Views']/i"), "'Views' option not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Save Options']/i"), "'Search' option not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Export']/i"), "'Export' option not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tblResultMain']"), "Calendar Grid not present.");

            Results.WriteStatus(test, "Pass", "Verified, Calendar Screen.");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Select Menu Option In Calendar View
        ///</summary>
        ///<param name="option">option to be selected</param>
        ///<param name="subOption">Sub Option to be selected</param>
        ///<returns></returns>
        public Calendar SelectMenuOptionInCalendarView(string option, string subOption = "", string view = "", string clientName = "")
        {
            string[] subOptionNameList = new string[1];
            string[] calendarDataByViewList = null;
            string[] calendarBasedOnViewList = null;

            if (clientName.ToLower().Contains("canada"))
            {
                calendarDataByViewList = new string[] { "Account", "Market", "Ad Type", "Manufacturer", "Category", "Brand", "Page Position", "Collection Type" };
                calendarBasedOnViewList = new string[] { "Drop Date", "Start Date", "End Date" };
            }
            else if (clientName.ToLower().Contains("australia"))
            {
                calendarDataByViewList = new string[] { "Retailer", "State", "Medium", "Manufacturer/Distributer", "Lead Category", "Category", "SubCategory", "Brand", "Page Position", "Collection Type" };
                calendarBasedOnViewList = new string[] { "Ad Date", "Media Start Date", "Media End Date" };
            }
            else
            {
                calendarDataByViewList = new string[] { "Retailer", "Market", "Ad Type", "Manufacturer", "Department", "Category Group", "Category", "Brand", "Page Position", "Media Type" };
                calendarBasedOnViewList = new string[] { "Ad Date", "Sale Start Date", "Sale End Date" };
            }

            if (option.ToLower().Contains("views"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Views']/i"), "'Views' option not present.");
                string[] specSubOptionNameList = new string[] { "VIEW OPTIONS", "Calendar Data By", "Calendar Based On", };
                Array.Resize(ref subOptionNameList, specSubOptionNameList.Length);
                Array.Copy(specSubOptionNameList, subOptionNameList, specSubOptionNameList.Length);

                if(!(clientName.ToLower().Contains("canada") || clientName.ToLower().Contains("australia")))
                {
                    Array.Resize(ref subOptionNameList, subOptionNameList.Length + 3);
                    subOptionNameList[3] = "VIEWS";
                    subOptionNameList[4] = "Normal";
                    subOptionNameList[5] = "Side By Side";
                }

                driver._click("xpath", "//navigation-menu//a[@title='Views']");
            }
            else if (option.ToLower().Contains("search"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Save Options']/i"), "'Search' option not present.");
                string[] specSubOptionNameList = new string[] { "Save Search" };
                Array.Resize(ref subOptionNameList, specSubOptionNameList.Length);
                Array.Copy(specSubOptionNameList, subOptionNameList, specSubOptionNameList.Length);

                driver._click("xpath", "//navigation-menu//a[@title='Save Options']");
            }
            else if (option.ToLower().Contains("export"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Export']/i"), "'Export' option not present.");
                string[] specSubOptionNameList = new string[] { "Excel", "Email Option" };
                Array.Resize(ref subOptionNameList, specSubOptionNameList.Length);
                Array.Copy(specSubOptionNameList, subOptionNameList, specSubOptionNameList.Length);

                driver._click("xpath", "//navigation-menu//a[@title='Export']");
            }

            Assert.IsTrue(driver._waitForElement("xpath", "//navigation-menu//div[contains(@class, 'open')]//li"), "'" + option + "' DDL not present.");
            IList<IWebElement> subOptionCollection = driver._findElements("xpath", "//navigation-menu//div[contains(@class, 'open')]//li");
            IWebElement subOptionEle = null;
            Actions action = new Actions(driver);
            foreach (string subName in subOptionNameList)
            {
                bool avail = false;
                foreach (IWebElement subEle in subOptionCollection)
                    if (subEle.Text.ToLower().Contains(subName.ToLower()))
                    {
                        avail = true;
                        if (subName.ToLower().Equals(subOption.ToLower()))
                            subOptionEle = subEle;
                        break;
                    }
                Assert.IsTrue(avail, "'" + subName + "' tab not found.");
            }

            if (subOption != "")
            {
                Assert.AreNotEqual(null, subOptionEle, "'" + subOption + "' Sub Option not found.");
                if (!subOptionEle.GetAttribute("class").Contains("active"))
                {
                    if (!(subOption.ToLower().Equals("calendar data by") || subOption.ToLower().Equals("calendar based on")))
                    {
                        if(!subOptionEle.GetAttribute("class").Contains("active"))
                            subOptionEle.Click();

                        Thread.Sleep(1000);
                        Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//navigation-menu//div[contains(@class, 'open')]//li"), "'" + option + "' DDL still present.");
                    }
                    else
                        action.MoveToElement(subOptionEle).MoveByOffset(2, 1).Perform();
                }
            }

            if(view != "")
            {
                IWebElement viewEle = null;
                if(subOption.ToLower().Equals("calendar data by"))
                {
                    Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'open')]//li[2]//li"), "'Calendar Data By' sub menu not present.");
                    IList<IWebElement> viewEleCollection = driver._findElements("xpath", "//div[contains(@class, 'open')]//li[2]//li");
                    foreach (string viewName in calendarDataByViewList)
                    {
                        bool avail = false;
                        foreach(IWebElement ele in viewEleCollection)
                            if (ele.Text.ToLower().Equals(viewName.ToLower()))
                            {
                                avail = true;
                                if (viewName.ToLower().Equals(view.ToLower()))
                                    viewEle = ele;
                                break;
                            }
                        Assert.IsTrue(avail, "'" + viewName + "' not found.");
                    }
                }
                else if (subOption.ToLower().Equals("calendar based on"))
                {
                    Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'open')]//li[3]//li"), "'Calendar Data By' sub menu not present.");
                    IList<IWebElement> viewEleCollection = driver._findElements("xpath", "//div[contains(@class, 'open')]//li[3]//li");
                    foreach (string viewName in calendarBasedOnViewList)
                    {
                        bool avail = false;
                        foreach (IWebElement ele in viewEleCollection)
                            if (ele.Text.ToLower().Equals(viewName.ToLower()))
                            {
                                avail = true;
                                if (viewName.ToLower().Equals(view.ToLower()))
                                    viewEle = ele;
                                break;
                            }
                        Assert.IsTrue(avail, "'" + viewName + "' not found.");
                    }
                }

                Assert.AreNotEqual(null, viewEle, "'" + view + "' not found.");
                if (!viewEle.GetAttribute("class").Contains("active"))
                    viewEle.Click();
                else
                    subOptionEle.Click();
                Thread.Sleep(3000);
            }

            Results.WriteStatus(test, "Pass", "Selected, '" + option + "' Menu Option In Calendar View");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Verify Calendar View Grid
        ///</summary>
        ///<param name="viewName">Name of the Calendar View Displayed</param>
        ///<param name="basedOn">Based On Criteria</param>
        ///<returns></returns>
        public Calendar VerifyCalendarViewGrid(string viewName = "Category", string basedOn = "Ad Date")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='tblResultMain']"), "Calendar View Grid not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='pageHeader']//td[not(@style)]"), "Calendar View Header not present.");

            if (viewName.ToLower().Equals("brand"))
            {
                Assert.AreEqual("calendar brand by (based on " + basedOn.ToLower() + ")", driver._getText("xpath", "//div[@id='pageHeader']//td[not(@style)]").ToLower(), "Calendar View Header text doesn't match.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='pageHeader']//input[@type='text']"), "Brand Price Field not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='pageHeader']//button"), "Brand Price DDL Button not present.");

                driver._click("xpath", "//div[@id='pageHeader']//button");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='tbldvBrandPrice']//td"), "Brand Price DDL not present.");
                IList<IWebElement> brandPriceDDLCol = driver._findElements("xpath", "//table[@id='tbldvBrandPrice']//td");

                Random rand = new Random();
                int x = rand.Next(0, brandPriceDDLCol.Count);
                string brandPrice = brandPriceDDLCol[x].Text;
                brandPriceDDLCol[x].Click();
                Thread.Sleep(3000);

                Assert.AreEqual(brandPrice, driver._getValue("xpath", "//div[@id='pageHeader']//input[@type='text']"), "Brand Price from DDL is not selected.");
            }
            else
                Assert.AreEqual("calendar view by " + viewName.ToLower() + " (based on " + basedOn.ToLower() + ")", driver._getText("xpath", "//div[@id='pageHeader']//td[not(@style)]").ToLower(), "Calendar View Header text doesn't match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//div[contains(@class, 'header')]"), "'Calendar header' not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//a[@title='<< Previous']"), "'Previous' button not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//a[@title='Next >>']"), "'Next' button not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]"), "'Month' field not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][2]"), "'Year' field not present");

            Assert.IsTrue(driver._isElementPresent("xpath", "//tr[@class='datepickerWeektitle']//span"), "'Week Day' names not present");
            IList<IWebElement> weekDayColl = driver._findElements("xpath", "//tr[@class='datepickerWeektitle']//span");
            string[] daysOfWeek = new string[] { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };

            foreach(string dayName in daysOfWeek)
            {
                bool avail = false;
                foreach(IWebElement weekDay in weekDayColl)
                    if (weekDay.Text.ToLower().Equals(dayName.ToLower()))
                    {
                        avail = true;
                        break;
                    }
                Assert.IsTrue(avail, "'" + dayName + "' not found");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class='ui-datepicker-other-month']"), "'Cells of past and next months' not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class=' ']"), "'Cells of current month' not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class=' ']/div[contains(@class, 'ui-state-default')]"), "'Dates' not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class=' ']/div[contains(@class, 'ui-datapicker-dvcontent')]"), "'Date View' not present");

            if(driver._isElementPresent("xpath", "//div[@id='calendar1']"))
                Results.WriteStatus(test, "Pass", "Verified, Calendar View Grid is displayed in Side By Side View.");
            else
                Results.WriteStatus(test, "Pass", "Verified, Calendar View Grid is displayed in Normal View.");

            Results.WriteStatus(test, "Pass", "Verified, Calendar View Grid.");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Capture Data From Calendar View Grid
        ///</summary>
        ///<returns></returns>
        public string[,] CaptureDataFromCalendarViewGrid()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//a[@title='Next >>']"), "'Next' button not present");

            string[] dateRow = new string[1];
            int i = 0, dataGridRows = 0;
            bool flag = true;

            while(flag)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]//div[@class='js_selectedText']"), "'Month' field not present");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][2]//div[@class='js_selectedText']"), "'Year' field not present");
                dateRow[i] = driver._getText("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]//div[@class='js_selectedText']");
                dateRow[i] = dateRow[i].Substring(0, 3) + "-" + driver._getText("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][2]//div[@class='js_selectedText']");
                dateRow[i] = dateRow[i].Substring(0, dateRow[i].Length - 2);
                IList<IWebElement> dateColl = driver._findElements("xpath", "//tr/td[@class=' ']/div[contains(@class, 'ui-state-default')]");
                int currRowCount = (int)(2 * (dateColl.Count / 7 + 1) + 3);
                dataGridRows = dataGridRows + currRowCount;

                if (driver._getAttributeValue("xpath", "//div[@id='calendar']//a[@title='Next >>']", "class").Contains("disabled"))
                    break;
                ++i;
                Array.Resize(ref dateRow, dateRow.Length + 1);
                driver._click("xpath", "//div[@id='calendar']//a[@title='Next >>']");
                Thread.Sleep(1000);
            } 

            foreach (string row in dateRow)
                Console.WriteLine(row);

            Console.WriteLine(dataGridRows);
            while(!driver._getAttributeValue("xpath", "//div[@id='calendar']//a[@title='<< Previous']", "class").Contains("disabled"))
            {
                driver._click("xpath", "//div[@id='calendar']//a[@title='<< Previous']");
                Thread.Sleep(1000);
            }

            string[,] dataGrid = new string[dataGridRows, 7];
            i = 0;
            int j = 0, k = 0;
            while (flag)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//tr[@class='datepickerWeektitle']//span"), "'Week Day' names not present");
                Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class='ui-datepicker-other-month']"), "'Cells of past and next months' not present");
                Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class=' ']"), "'Cells of current month' not present");
                Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class=' ']/div[contains(@class, 'ui-state-default')]"), "'Dates' not present");
                Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class=' ']/div[contains(@class, 'ui-datapicker-dvcontent')]"), "'Date View' not present");

                IList<IWebElement> weekDayColl = driver._findElements("xpath", "//tr[@class='datepickerWeektitle']//span");
                j = 0;
                dataGrid[i, j] = dateRow[k];
                ++i;
                int m = 0;
                for (; m < 7; m++, j++)
                    dataGrid[i, j] = weekDayColl[m].Text;
                ++i; m = 1; j = 0;
                while (driver._isElementPresent("xpath", "//div[@id='calendar']//tbody/tr[" + m + "]"))
                {
                    for (j = 0; j < 7; j++)
                        if (driver._isElementPresent("xpath", "//div[@id='calendar']//tbody/tr[" + m + "]//td[" + (j + 1) + "]/div[contains(@class, 'ui-state-default')]"))
                            dataGrid[i, j] = driver._getText("xpath", "//div[@id='calendar']//tbody/tr[" + m + "]//td[" + (j + 1) + "]/div[contains(@class, 'ui-state-default')]");
                    ++i;
                    for (j = 0; j < 7; j++)
                        if (driver._isElementPresent("xpath", "//div[@id='calendar']//tbody/tr[" + m + "]//td[" + (j + 1) + "]/div[contains(@class, 'dvcontent')]"))
                            dataGrid[i, j] = driver._getText("xpath", "//div[@id='calendar']//tbody/tr[" + m + "]//td[" + (j + 1) + "]/div[contains(@class, 'dvcontent')]");
                    ++i;
                    ++m;
                }

                if (driver._getAttributeValue("xpath", "//div[@id='calendar']//a[@title='Next >>']", "class").Contains("disabled"))
                    break;
                dataGrid[i, 0] = "";
                ++i; ++k;
                driver._click("xpath", "//div[@id='calendar']//a[@title='Next >>']");
                Thread.Sleep(1000);
            }

            for(i = 0; i < dataGrid.GetLength(0); i++)
            {
                for (j = 0; j < dataGrid.GetLength(1); j++)
                    Console.Write("(" + i + ", " + j + ")" + dataGrid[i, j] + "\t");
                Console.WriteLine();
            }

            Results.WriteStatus(test, "Pass", "Captured, Data From Calendar View Grid");
            return new string[0, 0];
        }

        ///<summary>
        ///Verify Data From Tabular Grid In Exported Excel File
        ///</summary>
        ///<param name="dataGrid">Data Captured from tabular grid</param>
        ///<param name="fileName">Name of Excel File</param>
        ///<returns></returns>
        public Calendar VerifyDataFromTabularGridInExportedExcelFile(string fileName, string[,] dataGrid)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            int rw = 0;
            int cl = 0;
            int flag = 1;
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
                if (xlWorkSheet.Name.Contains(" Online Tabular Report"))
                    break;
            }

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;





            //for (int i = 0, rCnt = 4; i < dataGrid.GetLength(0); i++, rCnt++)
            //{
            //    bool avail = false;
            //    Console.Write("\n(" + i + ", 0):\t" + dataGrid[i, 0] + " - " + (range.Cells[rCnt, 1] as Excel.Range).Text + "\t");
            //    if (dataGrid[i, 0].ToLower().Equals((range.Cells[rCnt, 1] as Excel.Range).Text.ToLower()))
            //    {
            //        flag = 1;
            //        for (int j = 1, cCnt = 2; j < dataGrid.GetLength(1); j++, cCnt++)
            //        {
            //            if (cCnt == 5)
            //                cCnt = 7;
            //            Console.Write("(" + i + ", " + j + "):\t" + dataGrid[i, j] + " - " + (range.Cells[rCnt, cCnt] as Excel.Range).Text + "\t");
            //            if (!(range.Cells[rCnt, cCnt] as Excel.Range).Text.ToLower().Contains(dataGrid[i, j].ToLower()))
            //            {
            //                --flag;
            //                break;
            //            }
            //        }
            //        if (flag > 0)
            //        {
            //            avail = true;
            //            break;
            //        }
            //    }
            //    Assert.IsTrue(avail, "Row '" + i + "' of Data Captured from Tabular Grid not found in Excel File.");
            //}

            Results.WriteStatus(test, "Pass", "Verified, Data from Calendar in exported file");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Verify Export Excel Popup
        ///</summary>
        ///<param name="popupVisible">Whether popup should be visible</param>
        ///<returns></returns>
        public Calendar VerifyExportExcelPopup(bool popupVisible = true)
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//table[@id='popupDiv1']"), "Export Excel Popup not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvTitle1']"), "Export Excel Popup header not present.");
                Assert.AreEqual("Options For Creating Your Calendar Report", driver._getText("xpath", "//div[@id='dvTitle1']"), "Export Excel Popup header text does not match.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='div_donotincludefeaturevisionhyperlink']"), "'Do not include hyperlinks' checkbox not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//td[text()='Do not include hyperlinks to the Numerator Promotions Intel site.']"), "'Do not include hyperlinks to the Numerator Promotions Intel site.' text not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='div_SaveDatatobeViewedlater']"), "'Save the Product Detail Records used' checkbox not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//td[text()='Save the Product Detail Records used to create your calendar report(s), and include links in the Report that will enable me to view all of these records and/or the associated images on the Numerator Promotions Intel site.']"), "'Save the Product Detail Records used to create your calendar report(s), and include links in the Report that will enable me to view all of these records and/or the associated images on the Numerator Promotions Intel site.' text not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='div_DoNotIncludeQueryParameter']"), "'Do not display query parameters above the data table.' checkbox not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//td[text()='Do not display query parameters above the data table.']"), "'Do not display query parameters above the data table.' text not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='div_SendReportinWinZipFormat']"), "'Deliver this report as a WinZip file' checkbox not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//td[text()='Deliver this report as a WinZip file to reduce file size (recommended).']"), "'Deliver this report as a WinZip file to reduce file size (recommended).' text not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='txtReportName']"), "'Use this as my Email subject line:' field not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//td[text()='Use this as my Email subject line:']"), "'Use this as my Email subject line:' not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='txtSubjectLine']"), "'Use this as my Email subject line:' field not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//td[text()='Use this name for my Report:']"), "'Use this name for my Report:' text not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//input[@value='Download Report']"), "'Download Report' not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[@value='Send Email With Attached Report']"), "'Send Email With Attached Report' button not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[@value='Cancel']"), "Cancel button not present.");
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//table[@id='popupDiv1']"), "Export Excel Popup still present.");
                Results.WriteStatus(test, "Pass", "Export Excel Pop is closed.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Export Excel Popup");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Select Checkbox or click button in Excel Popup
        ///</summary>
        ///<returns></returns>
        public Calendar SelectCheckboxOrClickButtonInExcelPopup(string checkboxOrButton)
        {
            if (checkboxOrButton.ToLower().Contains("download"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[@value='Download Report']"), "'Download Report' not present.");
                driver._click("xpath", "//input[@value='Download Report']");
                Results.WriteStatus(test, "Pass", "Clicked, Download Report button.");
            }
            else if (checkboxOrButton.ToLower().Contains("send email"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[@value='Send Email With Attached Report']"), "'Send Email With Attached Report' button not present.");
                driver._click("xpath", "//input[@value='Send Email With Attached Report']");
                Results.WriteStatus(test, "Pass", "Clicked, Send Email with Attached Report button.");
            }
            else if (checkboxOrButton.ToLower().Contains("cancel"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[@value='Cancel']"), "Cancel button not present.");
                driver._click("xpath", "//input[@value='Cancel']");
                Results.WriteStatus(test, "Pass", "Clicked, Cancel button.");
            }
            else if (checkboxOrButton.ToLower().Contains("do not include hyperlink"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='div_donotincludefeaturevisionhyperlink']"), "'Do not include hyperlinks' checkbox not present.");
                driver._click("xpath", "//div[@id='div_donotincludefeaturevisionhyperlink']");
                Results.WriteStatus(test, "Pass", "Selected, Do not include hyperlinks checkbox.");
            }
            else if (checkboxOrButton.ToLower().Contains("save the product detail records"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='div_SaveDatatobeViewedlater']"), "'Save the Product Detail Records used' checkbox not present.");
                driver._click("xpath", "//div[@id='div_SaveDatatobeViewedlater']");
                Results.WriteStatus(test, "Pass", "Selected, Save the Product Detail Records checkbox.");
            }
            else if (checkboxOrButton.ToLower().Contains("do not display query parameters"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='div_DoNotIncludeQueryParameter']"), "'Do not display query parameters above the data table.' checkbox not present.");
                driver._click("xpath", "//div[@id='div_DoNotIncludeQueryParameter']");
                Results.WriteStatus(test, "Pass", "Selected, Do not Display Query Parameters checkbox.");
            }
            else if (checkboxOrButton.ToLower().Contains("winzip"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='div_SendReportinWinZipFormat']"), "'Deliver this report as a WinZip file' checkbox not present.");
                driver._click("xpath", "//div[@id='div_SendReportinWinZipFormat']");
                Results.WriteStatus(test, "Pass", "Selected, Deliver this report as a WinZip file checkbox.");
            }

            return new Calendar(driver, test);
        }

        ///<summary>
        ///Enter Report Name And Email Subject Line in Export Excel Popup
        ///</summary>
        ///<returns></returns>
        public Calendar EnterReportNameAndEmailSubjectLineInExportExcelPopup(string reportName = "", string subjectLine = "")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='txtReportName']"), "'Use this as my Email subject line:' field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='txtSubjectLine']"), "'Use this as my Email subject line:' field not present.");

            if(reportName != "")
            {
                if (reportName.ToLower().Equals("random"))
                    reportName = "MyReport" + driver._randomString(3, true);

                driver._type("xpath", "//input[@id='txtReportName']", reportName);
                Thread.Sleep(1000);
                Results.WriteStatus(test, "Pass", "Entered, Report Name as '" + reportName + "'.");
            }

            if (subjectLine != "")
            {
                if (subjectLine.ToLower().Equals("random"))
                    subjectLine = "EmailSubject" + driver._randomString(3, true);

                driver._type("xpath", "//input[@id='txtSubjectLine']", subjectLine);
                Thread.Sleep(1000);
                Results.WriteStatus(test, "Pass", "Entered, Email Subject Line as '" + subjectLine + "'.");
            }


            Results.WriteStatus(test, "Pass", "Entered Report Name And/Or Email Subject Line in Export Excel Popup.");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Verify Send Email As Popup
        ///</summary>
        ///<param name="popupVisible">Whether popup should be visible</param>
        ///<param name="button">Button to be clicked</param>
        ///<returns></returns>
        public Calendar VerifySendEmailAsPopup(bool popupVisible = true, string button = "")
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "'Send Email As' popup not open.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']//h4"), "'Send Email As' popup header not present");
                Assert.AreEqual("Send Email As", driver._getText("xpath", "//div[@class='modal-content']//h4"), "'Send Email As' popup header not present");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//li[@heading]"), "Tabs not present.");
                string[] tabNamesList = new string[] { "Report", "Formatting Options", "Recipients", "Email Options" };
                IList<IWebElement> tabCollection = driver._findElements("xpath", "//div[@class='modal-body-filters']//li[@heading]");
                IWebElement recipientsTab = null;
                foreach(string tabName in tabNamesList)
                {
                    bool avail = false;
                    foreach(IWebElement tab in tabCollection)
                    {
                        if (tab.GetAttribute("heading").ToLower(). Equals(tabName.ToLower()))
                        {
                            avail = true;
                            if (tabName.ToLower().Equals("report"))
                                Assert.IsTrue(tab.GetAttribute("class").Contains("active"), "Report tab is not active.");
                            if (tabName.ToLower().Equals("recipients"))
                                recipientsTab = tab;
                            break;
                        }
                    }
                    Assert.IsTrue(avail, "'" + tabName + "' not found.");
                }

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[@id='leftcheckbox']"), "'Available Product Detail Report and Summary Templates' box not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[@id='rightcheckbox']"), "'Available Report/Summary Groups' box not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//label[@ng-if='data.IncludeCalendar.Visible']"), "'Include Calendar' Checkbox Label not present.");
                IWebElement checkbox = driver.FindElement(By.XPath("//input[@ng-model='data.IncludeCalendar.Checked']"));
                Assert.AreNotEqual(null, checkbox.GetAttribute("checked"), "'Include Calendar' Checkbox is not checked by default.");

                recipientsTab.Click();
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@class='modal-body-filters']//li[@heading='Recipients']", "class").Contains("active"), "Recipients tab is not selected.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[@id='leftcheckboxEmail']//li"), "'Available Recipients' box not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[@id='rightcheckboxEmail']//li"), "'Selected Recipients' box not present.");
                checkbox = driver.FindElement(By.XPath("//div[@id='rightcheckboxEmail']//li//input"));
                Assert.AreNotEqual(null, checkbox.GetAttribute("checked"), "Email Recipient not selected.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//button[text()='Email As Attachment']"), "'Email As Attachment' button not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//button[text()='Email with Download Link']"), "'Email with Download Link' button not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//button[text()='Cancel']"), "'Cancel' button not present.");

                if (button.ToLower().Contains("attachment"))
                    driver._click("xpath", "//button[text()='Email As Attachment']");
                else if (button.ToLower().Contains("download"))
                    driver._click("xpath", "//button[text()='Email with Download Link']");
                else
                    driver._click("xpath", "//button[text()='Cancel']");
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "'Send Email As' popup is still open.");
                Results.WriteStatus(test, "Pass", "Verified, 'Send Email As' popup is closed.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Send Email As Popup.");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Select Month From DDL in Calendar
        ///</summary>
        ///<param name="month">Month to be selected</param>
        ///<returns></returns>
        public Calendar SelectMonthFromDDLInCalendar(string month = "April")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]"), "'Month' field not present");
            driver._click("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]//li/span"), "Month DDL not present.");
            IList<IWebElement> monthDDLColl = driver._findElements("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]//li/span");

            bool avail = false;
            foreach(IWebElement monthEle in monthDDLColl)
            {
                if (monthEle.Text.ToLower().Contains(month.ToLower()))
                {
                    avail = true;
                    monthEle.Click();
                    break;
                }
            }
            Assert.IsTrue(avail, "'" + month + "' not found in month DDL");

            Thread.Sleep(2000);
            Assert.IsTrue(month.ToLower().Contains(driver._getText("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]/div").ToLower()), "'" + month + "' not selected.");

            Results.WriteStatus(test, "Pass", "Selected, '" + month + "' Month From DDL in Calendar.");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Select Year From DDL in Calendar
        ///</summary>
        ///<param name="year">Year to be selected</param>
        ///<returns></returns>
        public Calendar SelectYearFromDDLInCalendar(string year = "2020")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][2]"), "'Year' field not present");
            driver._click("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][2]");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][2]//li/span"), "Year DDL not present.");
            IList<IWebElement> yearDDLColl = driver._findElements("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][2]//li/span");

            bool avail = false;
            foreach (IWebElement yearEle in yearDDLColl)
            {
                if (yearEle.Text.ToLower().Contains(year.ToLower()))
                {
                    avail = true;
                    yearEle.Click();
                    break;
                }
            }
            Assert.IsTrue(avail, "'" + year + "' not found in Year DDL");

            Thread.Sleep(2000);
            Assert.IsTrue(year.ToLower().Contains(driver._getText("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][2]/div").ToLower()), "'" + year + "' not selected.");

            Results.WriteStatus(test, "Pass", "Selected, '" + year + "' Year From DDL in Calendar.");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Verify Calendar Navigation
        ///</summary>
        ///<param name="next">Whether to navigate forward or backward</param>
        ///<returns></returns>
        public Calendar VerifyCalendarNavigation(bool next = true)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//a[@title='<< Previous']"), "'Previous' button not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='calendar']//a[@title='Next >>']"), "'Next' button not present");

            string[] monthsList = new string[] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            string currentMonth = driver._getText("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]/div");

            int index = -1;
            for(int i = 0; i < monthsList.Length; i++)
                if (monthsList[i].ToLower().Equals(currentMonth.ToLower()))
                {
                    index = i;
                    break;
                }
            Assert.Greater(index, -1, "'" + currentMonth + "' is not from 12 Months of Calendar.");

            string newMonth = "";

            if (next)
            {
                Assert.IsFalse(driver._getAttributeValue("xpath", "//div[@id='calendar']//a[@title='Next >>']", "class").Contains("disabled"), "Next Button is disabled.");
                driver._click("xpath", "//div[@id='calendar']//a[@title='Next >>']");
                Thread.Sleep(2000);
                newMonth = driver._getText("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]/div");
                Assert.AreEqual(newMonth.ToLower(), monthsList[index + 1].ToLower(), "Previous button was not clicked.");
                Results.WriteStatus(test, "Pass", "Verified, Next Navigation button on calendar");
            }
            else
            {
                Assert.IsFalse(driver._getAttributeValue("xpath", "//div[@id='calendar']//a[@title='<< Previous']", "class").Contains("disabled"), "Previous Button is disabled.");
                driver._click("xpath", "//div[@id='calendar']//a[@title='<< Previous']");
                Thread.Sleep(2000);
                newMonth = driver._getText("xpath", "//div[@id='calendar']//div[@class='js_dropdown'][1]/div");
                Assert.AreEqual(newMonth.ToLower(), monthsList[index - 1].ToLower(), "Previous button was not clicked.");
                Results.WriteStatus(test, "Pass", "Verified, Previous Navigation button on calendar");
            }

            Results.WriteStatus(test, "Pass", "Verified, Calendar Navigation");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Verify Carousel From Calendar Screen
        ///</summary>
        ///<returns></returns>
        public Calendar VerifyCarouselFromCalendarScreen()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class=' ']/div[contains(@class, 'ui-datapicker-dvcontent')]//a"), "Date Link not present.");
            IList<IWebElement> dateLinkColl = driver._findElements("xpath", "//tr/td[@class=' ']/div[contains(@class, 'ui-datapicker-dvcontent')]//a");

            dateLinkColl[0].Click();
            Thread.Sleep(3000);
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li"), "Carousel not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='adBlockCarousel']//div[@class='fa fa-share-square-o']"), "Export Icon not present.");

            IList<IWebElement> carouselCard = driver._findElements("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li");
            driver._scrollintoViewElement("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]");
            Actions action = new Actions(driver);
            action.MoveToElement(carouselCard[0]).MoveByOffset(4, 4).Perform();

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//p/a[contains(text(), 'View Ad')]"), "View Ad not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//p/a[contains(text(), 'Detail')]"), "Detail not present.");
            
            Results.WriteStatus(test, "Pass", "Verified, Carousel From Calendar Screen");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Open View Ad Or Detail Popup
        ///</summary>
        ///<param name="openDetail">Whether to open Detail Tab</param>
        ///<returns></returns>
        public Calendar OpenViewAdOrDetailPopup(bool openDetail)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li"), "Carousel not present.");

            IList<IWebElement> carouselCard = driver._findElements("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li");
            driver._scrollintoViewElement("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]");
            Actions action = new Actions(driver);
            action.MoveToElement(carouselCard[0]).MoveByOffset(4, 4).Perform();

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//p/a[contains(text(), 'View Ad')]"), "View Ad not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//p/a[contains(text(), 'Detail')]"), "Detail not present.");


            if (openDetail)
            {
                driver._click("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//p/a[contains(text(), 'Detail')]");

                Results.WriteStatus(test, "Pass", "Opened, 'Detail' Popup from Calendar View.");
            }
            else
            {
                driver._click("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//p/a[contains(text(), 'View Ad')]");

                Results.WriteStatus(test, "Pass", "Opened, 'View Ad' Popup from Calendar View.");
            }

            return new Calendar(driver, test);
        }

        ///<summary>
        ///Verify Export As Excel Option From Carousel
        ///</summary>
        ///<returns></returns>
        public Calendar VerifyExportAsExcelOptionFromCarousel()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='adBlockCarousel']//div[@class='fa fa-share-square-o']"), "Export Icon not present.");
            driver._click("xpath", "//div[@id='adBlockCarousel']//div[@class='fa fa-share-square-o']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='adBlockCarousel']//ul[@class='ulChartOption']/li[text()='Export as EXCEL']"), "'Export as EXCEL' option not present");
            driver._click("xpath", "//div[@id='adBlockCarousel']//ul[@class='ulChartOption']/li[text()='Export as EXCEL']");
            Thread.Sleep(10000);

            Results.WriteStatus(test, "Pass", "Verified, Export As Excel Option From Carousel");
            return new Calendar(driver, test);
        }

        ///<summary>
        ///Verify Navigation On Calendar Carousel
        ///</summary>
        ///<returns></returns>
        public Calendar VerifyNavigationOnCalendarCarousel()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//tr/td[@class=' ']/div[contains(@class, 'ui-datapicker-dvcontent')]//a"), "Date Link not present.");
            IList<IWebElement> dateLinkColl = driver._findElements("xpath", "//tr/td[@class=' ']/div[contains(@class, 'ui-datapicker-dvcontent')]//a");

            foreach(IWebElement dateLink in dateLinkColl)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", dateLink);
                string dateLinkText = dateLink.Text;
                int dateLinkNumber = 0;
                Assert.IsTrue(int.TryParse(dateLinkText, out dateLinkNumber), "Date Link Text Could not be converted to int.");
                if (dateLinkNumber > 6)
                {
                    dateLink.Click();
                    break;
                }
            }

            Thread.Sleep(3000);
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='left carousel-control']"), "Left Navigation Button not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='right carousel-control']"), "Left Navigation Button not present.");

            string firstCard = driver._getText("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//div[@class='col-xs-8 tLabel']/p");
            driver._click("xpath", "//div[@class='right carousel-control']");
            Thread.Sleep(5000);
            string newCard = driver._getText("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//div[@class='col-xs-8 tLabel']/p");
            Assert.AreNotEqual(firstCard, newCard, "Next Navigation button did not work.");

            driver._click("xpath", "//div[@class='left carousel-control']");
            Thread.Sleep(5000);
            string prevCard = driver._getText("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//div[@class='col-xs-8 tLabel']/p");
            Assert.AreNotEqual(newCard, prevCard, "Previous Navigation button did not work.");
            Assert.AreEqual(firstCard, prevCard, "Previous Navigation button did not work.");

            Results.WriteStatus(test, "Pass", "Verified, Navigation On Calendar Carousel");
            return new Calendar(driver, test);
        }

        #endregion
    }
}
