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
    public class ManufacturerComparison
    {
        #region Private Variables

        private IWebDriver manufacturerComparison;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public ManufacturerComparison(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.manufacturerComparison = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.manufacturerComparison; }
            set { this.manufacturerComparison = value; }
        }

        ///<summary>
        ///Verify Ad Sharing And Exclusivity Page
        ///</summary>
        ///<returns></returns>
        public ManufacturerComparison VerifyManufacturerComparisonPage(bool fromPricingPromotions = true, string searchName = "")
        {
            if (fromPricingPromotions)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//h1[text()='Pricing & Promotions']"), "'Pricing & Promotions' Screen not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='title cursorpointer']/a[text()='Manufacturer Comparison']"), "'Manufacturer Comparison' link not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-media']/i[contains(@class, 'manufacturer-comparison')]"), "'Manufacturer Comparison' icon not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-content']/div[contains(@class, 'submenudescription')]"), "'Manufacturer Comparison' description not present.");

                driver._click("xpath", "//div[@class='title cursorpointer']/a[text()='Manufacturer Comparison']");
                Thread.Sleep(1000);
            }

            driver._waitForElementToBeHidden("xpath", "//div[@class='ProcessLoader'], 45");

            Assert.IsTrue(driver._waitForElement("xpath", "//img[@src='/Images/isc-NumeratorLogo.png']", 15), "Numerator Logo not found on Home Page.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='report-content']/h1"), "Madlib Search header text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[contains(@class, 'filler')]"), "Madlib Prompt Summary text not present.");

            if (searchName != "")
                Assert.IsTrue(driver._getText("xpath", "//div[@class='report-content']/h1").ToLower().Contains(searchName.ToLower()), "Search Name '" + searchName + "' does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='1']"), "Madlib Search Parameter 'Any Product' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='2']"), "Madlib Search Parameter 'Any Account' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='3']"), "Madlib Search Parameter 'Any Market' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='4']"), "Madlib Search Parameter 'Any Date' not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvComparisonGridContainerHeaderTitle']//li/span"), "'Manufacturer Comparison' Header text not present.");
            Assert.AreEqual("Manufacturer Comparison", driver._getText("xpath", "//div[@id='dvComparisonGridContainerHeaderTitle']//li/span"), "'Manufacturer Comparison' Header text doesn't match");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvComparisonGridContainerHeader']//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvComparisonGridContainerHeader']//div[@title='Export']"), "'Export' Icon not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvComparisonGridContainer']//div[contains(@class, 'GridCont')]"), "'Grid' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvComparisonGridContainer']//div[contains(@class, 'row ng-scope')]"), "'Rows' in Grid not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvComparisonGridContainer']//div[@title]"), "'Column and Row Headers' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'row ng-scope')]/div"), "'Cells' in Grid not present.");

            string[] rowHeaderList = new string[] { "# Products", "# Pages", "# of Unique ads", "Average Unit Price", "% of Promoted Products Pictured" };
            IList<IWebElement> headerColl = driver._findElements("xpath", "//div[@id='dvComparisonGridContainer']//div[@title]");

            foreach(string rowHeader in rowHeaderList)
            {
                bool avail = false;
                foreach(IWebElement header in headerColl)
                    if (header.GetAttribute("title").ToLower().Equals(rowHeader.ToLower()))
                    {
                        avail = true;
                        break;
                    }
                Assert.IsTrue(avail, "'" + rowHeader + "' not found.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Manufacturer Comparison Page");
            return new ManufacturerComparison(driver, test);
        }

        ///<summary>
        ///Verify Manufacturer Comparison Help Icon
        ///</summary>
        ///<returns></returns>
        public ManufacturerComparison VerifyManufacturerComparisonPageHelpIcon()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvComparisonGridContainerHeader']//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            driver._scrollintoViewElement("xpath", "//div[@id='dvComparisonGridContainerHeader']//div[@class='fa fa-question homeChartSearch']");
            driver.MouseHoverUsingElement("xpath", "//div[@id='dvComparisonGridContainerHeader']//div[@class='fa fa-question homeChartSearch']");

            string capturedHelpIconText = driver._getAttributeValue("xpath", "//div[@id='dvComparisonGridContainerHeader']//div[@onmouseout='HideTooltip();']", "onmouseover");
            string expectedHelpIconText1 = "How do I compare against other manufacturers? What manufacturers are retailers promoting most often in a category?";
            string expectedHelpIconText2 = "Provides a quick manufacturer scorecard with side-by-side comparisons. Compare key promotional metrics across three manufacturers at a time. Click on the manufacturer name / dropdown in the header to change the manufacturer. Green, yellow and red coloring indicate which manufacturer won, tied, and lost respectively.Drill into # Products to pull up the detail data or thumbnails.";

            Console.WriteLine(capturedHelpIconText);
            Assert.IsTrue(capturedHelpIconText.ToLower().Contains(expectedHelpIconText1.ToLower()) && capturedHelpIconText.ToLower().Contains(expectedHelpIconText2.ToLower()), "Help Icon tooltip text did not match.");

            Results.WriteStatus(test, "Pass", "Verified, Help Icon for 'Manufacturer Comparison' Page.");
            return new ManufacturerComparison(driver, test);
        }

        ///<summary>
        ///Verify Column DDL On Manufacturer Comparison Grid
        ///</summary>
        ///<returns></returns>
        public ManufacturerComparison VerifyColumnDDLOnManufacturerComparisonGrid()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'GridCont')]//button"), "'DDL' buttons not present on Manufacturer Comparison Grid");
            IList<IWebElement> buttonColl = driver._findElements("xpath", "//div[contains(@class, 'GridCont')]//button");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'row ng-scope')]//div[3]/*[contains(@class, 'countText')]"), "Cells in 2nd Column not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'row ng-scope')]//div[4]/*[contains(@class, 'countText')]"), "Cells in 2nd Column not present.");

            IList<IWebElement> cellCollection = driver._findElements("xpath", "//div[contains(@class, 'row ng-scope')]//div[3]/*[contains(@class, 'countText')]");
            string[] prevCellTextList = new string[cellCollection.Count];

            for (int i = 0; i < prevCellTextList.Length; i++)
                prevCellTextList[i] = cellCollection[i].Text;

            buttonColl[0].Click();
            Assert.IsTrue(driver._waitForElement("xpath", "//ul[@aria-labelledby='dropdownMenu2']//a"), "'DDL' not present.");
            IList<IWebElement> ddlCollection = driver._findElements("xpath", "//ul[@aria-labelledby='dropdownMenu2']//a");

            Random rand = new Random();
            int x = rand.Next(0, ddlCollection.Count);

            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", ddlCollection[x]);
            string newValue = ddlCollection[x].Text;
            ddlCollection[x].Click();

            Thread.Sleep(4000);
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'row ng-scope')]//div[3]/*[contains(@class, 'countText')]"), "Cells in 2nd Column not present.");
            cellCollection = driver._findElements("xpath", "//div[contains(@class, 'row ng-scope')]//div[3]/*[contains(@class, 'countText')]");
            string[] newCellTextList = new string[cellCollection.Count];

            Assert.IsTrue(cellCollection[0].Text.Contains(newValue), "New Value '" + newValue + "' from DDL not selected.");

            for (int i = 0; i < newCellTextList.Length; i++)
                newCellTextList[i] = cellCollection[i].Text;

            bool avail = false;
            for (int i = 0; i < newCellTextList.Length - 1; i++)
            {
                Console.WriteLine("Previous Value : " + prevCellTextList[i] + "\t;\tNew Value : " + newCellTextList[i]);
                if (prevCellTextList[i].Equals(newCellTextList[i]))
                {
                    avail = true;
                    break;
                }
            }
            Assert.IsFalse(avail, "'" + newValue + "' not applied.");

            Results.WriteStatus(test, "Pass", "Verified, Column DDL On Manufacturer Comparison Grid");
            return new ManufacturerComparison(driver, test);
        }

        ///<summary>
        ///Verify Detail Data on Manufacturer Comparison Page
        ///</summary>
        ///<returns></returns>
        public ManufacturerComparison VerifyDetailDataOnManufacturerComparisonPage()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'row ng-scope')]//a[contains(@class, 'countText')]"), "'# Product' rows not present.");
            IList<IWebElement> productRowColl = driver._findElements("xpath", "//div[contains(@class, 'row ng-scope')]//a[contains(@class, 'countText')]");

            productRowColl[0].Click();
            Thread.Sleep(2000);

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='drilledDataWidgetContainer']"), "'Detail Data' section not present.");
            driver._scrollintoViewElement("xpath", "//div[@id='drilledDataWidgetContainer']");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'detailHeader')]//span[contains(@class, 'viewModeText')]"), "'Detail Data' Header not present.");
            Assert.AreEqual("Detail Data", driver._getText("xpath", "//div[contains(@class, 'detailHeader')]//span[contains(@class, 'viewModeText')]"), "'Detail Data' Header text does not match.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'detailHeader')]//span[contains(@class, 'extraText')]"), "'Client Name' not present.");
            Assert.IsTrue(driver._getText("xpath", "//div[contains(@class, 'detailHeader')]//span[contains(@class, 'extraText')]").Contains("Procter & Gamble"), "'Procter & Gamble' client name does not match.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'chkLbl') and text()='Detail Data']"), "'Detail Data' Radio button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'chkLbl') and text()='Promoted Product Images']"), "'Promoted Product Images' Radio button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'chkLbl') and text()='Page Images']"), "'Page Images' Radio button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'detailHeader')]//div[@class='CustChartfilter-header']"), "'Export' icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'pageSizeGrp')]/button"), "'Records per Page' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//ul[contains(@class, 'pagination')]//a"), "'Page Navigation' not present.");

            Results.WriteStatus(test, "Pass", "Verified, Detail Data on Manufacturer Comparison Page");
            return new ManufacturerComparison(driver, test);
        }

        ///<summary>
        ///Verify Radio Buttons On Detail Data Section
        ///</summary>
        ///<returns></returns>
        public ManufacturerComparison VerifyRadioButtonsOnDetailDataSection(string radioButton = "Detail Data")
        {
            if (radioButton.ToLower().Equals("detail data"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'chkLbl') and text()='Detail Data']"), "'Detail Data' Radio button not present.");
                driver._click("xpath", "//span[contains(@class, 'chkLbl') and text()='Detail Data']");
                Thread.Sleep(2000);

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@ag-grid]"), "'Detail Data' Table not present.");

            }
            else if (radioButton.ToLower().Equals("promoted product images"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'chkLbl') and text()='Promoted Product Images']"), "'Promoted Product Images' Radio button not present.");
                driver._click("xpath", "//span[contains(@class, 'chkLbl') and text()='Promoted Product Images']");
                Thread.Sleep(2000);

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='imageView']/div"), "'Promoted Product Images' not present.");
            }
            else if (radioButton.ToLower().Equals("page images"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'chkLbl') and text()='Page Images']"), "'Page Images' Radio button not present.");
                driver._click("xpath", "//span[contains(@class, 'chkLbl') and text()='Page Images']");
                Thread.Sleep(2000);

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='imageView']/div"), "'Page Images' not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'aditem-details-1prod')]//div[text()='Circular' or text()='Bonus Pages']"), "Images are not of 'Page Images'.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Radio Buttons On Detail Data Section");
            return new ManufacturerComparison(driver, test);
        }

        ///<summary>
        ///Verify Export As Excel Option From Detail Data Section
        ///</summary>
        ///<returns></returns>
        public ManufacturerComparison VerifyExportAsExcelOptionFromDetailDataSection()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'detailHeader')]//div[@class='fa fa-share-square-o']"), "Export Icon not present.");
            driver._click("xpath", "//div[contains(@class, 'detailHeader')]//div[@class='fa fa-share-square-o']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'detailHeader')]//ul[@class='ulChartOption']/li[text()='Export as EXCEL']"), "'Export as EXCEL' option not present");
            driver._click("xpath", "//div[contains(@class, 'detailHeader')]//ul[@class='ulChartOption']/li[text()='Export as EXCEL']");
            Thread.Sleep(10000);

            Results.WriteStatus(test, "Pass", "Verified, Export As Excel Option From Detail Data Section");
            return new ManufacturerComparison(driver, test);
        }

        ///<summary>
        ///Capture Data From Detail Data Table
        ///</summary>
        ///<returns></returns>
        public string[,] CaptureDataFromDetailDataTable()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'ag-header-cell-label')]//span[@id='agText']"), "Header Cells not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@row]"), "'Data Row' not present.");

            string[] columnHeaderList = new string[] { "Ad Date", "Retailer", "Market", "Media Type", "Category", "Brand", "Product Description",
                                        "Product Size", "Ad Price", "Offer"};
            string[] colidList = new string[] { "4", "1", "5", "210", "14", "15", "126", "142", "32", "178" };
            Actions action = new Actions(driver);

            string[,] dataGrid = new string[10, 10];
            for (int i = 0; i < dataGrid.GetLength(1); i++)
                dataGrid[0, i] = columnHeaderList[i];

            for(int j = 0; j < dataGrid.GetLength(1); j++)
            {
                for(int i = 1; i < dataGrid.GetLength(0); i++)
                {
                    driver._scrollintoViewElement("xpath", "//div[@row=" + (i - 1) + "]//div[@colid='" + colidList[j] + "']");
                    dataGrid[i, j] = driver._getText("xpath", "//div[@row=" + (i - 1) + "]//div[@colid='" + colidList[j] + "']");
                }
                if(j == 4)
                {
                    driver._click("xpath", "//div[@row=8]//div[@colid='" + colidList[j] + "']");
                    for (int k = 0; k < 5; k++)
                        action.SendKeys(Keys.ArrowRight);
                }
            }

            for (int i = 1; i < dataGrid.GetLength(1); i++)
            {
                while (dataGrid[i, 0].IndexOf("/") > -1)
                    dataGrid[i, 0] = dataGrid[i, 0].Replace("/", "-");
            }

            Results.WriteStatus(test, "Pass", "Captured, Data From Detail Data Table");
            return dataGrid;
        }
        
        ///<summary>
        ///Verify Data From Tabular Grid In Exported Excel File
        ///</summary>
        ///<param name="dataGrid">Data Captured from tabular grid</param>
        ///<param name="fileName">Name of Excel File</param>
        ///<returns></returns>
        public ManufacturerComparison VerifyDataFromTabularGridInExportedExcelFile(string fileName, string[,] dataGrid)
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
                if (xlWorkSheet.Name.Contains("Promoted Products"))
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
                    Console.WriteLine("Excel : " + (range.Cells[rCnt, cCnt] as Excel.Range).Text + " : dataGrid : " + dataGrid[i,j]);
                    string temp = (range.Cells[rCnt, cCnt] as Excel.Range).Text;
                    while (temp.IndexOf("  ") > -1)
                        temp = temp.Replace("  ", " ");
                    Assert.IsTrue(temp.ToLower().Contains(dataGrid[i, j].ToLower()), "Data Incorrect for (" + i + ", " + j + ")");
                }
            }

            Results.WriteStatus(test, "Pass", "Verified, Data from Calendar in exported file");
            return new ManufacturerComparison(driver, test);
        }

        ///<summary>
        ///Verify Items Per Page Functionality
        ///</summary>
        ///<param name="noOfItems">No of Items per page to be selected</param>
        ///<returns></returns>
        public ManufacturerComparison VerifyItemsPerPageFunctionality(string noOfItems = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'pageSize')]//button/a"), "Items per page buttons not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='imageView']/div"), "Records not present.");
            IList<IWebElement> recordsColl = driver._findElements("xpath", "//div[@id='imageView']/div");
            string selectedItemsNum = driver._getText("xpath", "//div[contains(@class, 'pageSize')]//button[contains(@class, 'active')]/a");
            Assert.LessOrEqual(recordsColl.Count.ToString(), selectedItemsNum, "Selected Items per page do not match the displayed no. of items.");

            if (selectedItemsNum.Equals(noOfItems))
                Results.WriteStatus(test, "Pass", "'" + noOfItems + "' Items per page is already selected.");
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'pageSize')]//button[@btn-radio='" + noOfItems + "']"), "'" + noOfItems + "' Items per page button not present.");
                driver._click("xpath", "//div[contains(@class, 'pageSize')]//button[@btn-radio='" + noOfItems + "']");
                homePage.VerifyHomePage();
                Thread.Sleep(2000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[contains(@class, 'pageSize')]//button[@btn-radio='" + noOfItems + "']", "class").Contains("active"), "'" + noOfItems + "' Items per page button was not selected.");
                recordsColl = driver._findElements("xpath", "//div[@id='imageView']/div");
                Assert.LessOrEqual(recordsColl.Count.ToString(), noOfItems, "Selected Items per page do not match the displayed no. of items.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Items per page functionality by selecting to '" + noOfItems + "' Items per page.");
            return new ManufacturerComparison(driver, test);
        }

        ///<summary>
        ///Verify Filter Functionality On Detail Data Table
        ///</summary>
        ///<returns></returns>
        public ManufacturerComparison VerifyFilterFunctionalityOnDetailDataTable()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'ag-header-viewport')]//div[@colid='14']"), "'Category Column Header' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@colid='14']//span[@id='agMenu']"), "'Filter' icon not present on Category Header");
            driver._click("xpath", "//div[@colid='14']//span[@id='agMenu']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@id, 'agGridFilterTabBody')]"), "Filter Menu not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'agGridFilterTabBody')]//input[@type='text']"), "'Search' Text bar not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'agGridFilterTabBody')]//span[text()='(Select All)']"), "'Select All' switch not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'agGridFilterTabBody')]//div[@class='ag-virtual-list-item']//span"), "Filter list not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'ag-tab-body')]//button[text()='Apply Filter']"), "Apply Filter button not present.");

            IWebElement selectAllSwitch = driver.FindElement(By.XPath("//div[contains(@class, 'ag-tab-body')]//input[contains(@id, 'SelectAll')]"));
            Thread.Sleep(2000);
            driver._click("xpath", "//div[contains(@id, 'agGridFilterTabBody')]//span[text()='(Select All)']");
            Thread.Sleep(2000);
            Assert.IsTrue(selectAllSwitch.Selected, "Select All switch is not ON");
            IList<IWebElement> checkboxColl = driver._findElements("xpath", "//div[contains(@class, 'ag-tab-body')]//input[@idval]");

            bool avail = true;
            foreach(IWebElement checkbox in checkboxColl)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", checkbox);
                if (!checkbox.Selected)
                {
                    avail = false;
                    break;
                }
            }
            Assert.IsTrue(avail, "All filter options are not selected.");

            driver._click("xpath", "//div[contains(@id, 'agGridFilterTabBody')]//span[text()='(Select All)']");
            Thread.Sleep(2000);
            Assert.IsFalse(selectAllSwitch.Selected, "Select All switch is not OFF");
            checkboxColl = driver._findElements("xpath", "//div[contains(@class, 'ag-tab-body')]//input[@idval]");

            avail = true;
            foreach (IWebElement checkbox in checkboxColl)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", checkbox);
                if (checkbox.Selected)
                {
                    avail = false;
                    break;
                }
            }
            Assert.IsTrue(avail, "All filter options are deselected.");

            string filterValue = checkboxColl[0].GetAttribute("val");
            driver._click("xpath", "//div[contains(@id, 'agGridFilter')]//div[@class='ag-virtual-list-item'][1]//span");
            Assert.IsTrue(checkboxColl[0].Selected, "'" + filterValue + "' is not selected.");
            driver._click("xpath", "//div[contains(@class, 'ag-tab-body')]//button[text()='Apply Filter']");
            Thread.Sleep(2000);

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]"), "Category Column cells not present.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]");

            avail = true;
            foreach (IWebElement cell in cellCollection)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cell);
                if (!cell.Text.ToLower().Equals(filterValue.ToLower()))
                {
                    avail = false;
                    break;
                }
            }
            Assert.IsTrue(avail, "'" + filterValue + "' was not applied successfully.");

            driver._scrollintoViewElement("xpath", "//div[@colid='14']//span[@id='agMenu']");
            driver._click("xpath", "//div[@colid='14']//span[@id='agMenu']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@id, 'agGridFilterTabBody')]"), "Filter Menu not present.");
            driver._click("xpath", "//div[contains(@class, 'ag-tab-body')]//button[text()='Remove Filter']");

            Thread.Sleep(2000);

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]"), "Category Column cells not present.");
            cellCollection = driver._findElements("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]");

            avail = false;
            foreach (IWebElement cell in cellCollection)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cell);
                if (!cell.Text.ToLower().Equals(filterValue.ToLower()))
                {
                    avail = true;
                    break;
                }
            }
            Assert.IsTrue(avail, "'" + filterValue + "' was not removed successfully.");

            Results.WriteStatus(test, "Pass", "Verified, Filter Functionality On Detail Data Table");
            return new ManufacturerComparison(driver, test);
        }

        ///<summary>
        ///Verify Sort Fucntionality On Detail Data Section
        ///</summary>
        ///<return></return>
        public ManufacturerComparison VerifySortFunctionalityOnDetailDataSection()
        {
            Thread.Sleep(3000);
            driver._click("xpath", "//div[contains(@class, 'pageSize')]//button[@btn-radio='20']");
            Thread.Sleep(3000);
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'ag-header-viewport')]//div[@colid='14']"), "'Category Column Header' not present.");
            driver._click("xpath", "//div[@row=0]//div[@colid='14']");
            Actions action = new Actions(driver);
            action.SendKeys(Keys.ArrowRight).Perform();

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@colid='14']//span[contains(@id, 'Sort') and not(contains(@class, 'hidden'))]"), "'Sort' icon not present on Category Header");
            Thread.Sleep(2000);
            driver._click("xpath", "//div[@colid='14']//span[@id='agText']");
            Thread.Sleep(3000);
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]"), "Category Column cells not present.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]");

            string[] cellValueList = new string[cellCollection.Count];
            for(int i = 0; i < cellCollection.Count; i++)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cellCollection[i]);
                cellValueList[i] = cellCollection[i].Text;
            }

            string[] cellValueListCopy = new string[cellValueList.Length];
            Array.Copy(cellValueList, cellValueListCopy, cellValueList.Length);

            if(driver._getAttributeValue("xpath", "//div[@colid='14']//span[contains(@id, 'Sort') and not(contains(@class, 'hidden'))]", "id").Contains("Asc"))
            {
                Array.Sort(cellValueListCopy);
                Assert.IsTrue(cellValueList.SequenceEqual(cellValueListCopy), "Category Column is not sorted in Ascending order.");
                Results.WriteStatus(test, "Pass", "Verified, Category Column is sorted in ascending order.");

                driver._click("xpath", "//div[@colid='14']//span[@id='agText']");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@colid='14']//span[contains(@id, 'Sort') and not(contains(@class, 'hidden'))]", "id").Contains("Desc"), "Category column is not sorted in descending order.");
                driver._scrollintoViewElement("xpath", "//div[@row=0]//div[@colid='14']");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]"), "Category Column cells not present.");
                cellCollection = driver._findElements("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]");

                cellValueList = new string[cellCollection.Count];
                for (int i = 0; i < cellCollection.Count; i++)
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cellCollection[i]);
                    cellValueList[i] = cellCollection[i].Text;
                }

                cellValueListCopy = new string[cellValueList.Length];
                Array.Copy(cellValueList, cellValueListCopy, cellValueList.Length);
                Array.Sort(cellValueListCopy);
                Array.Reverse(cellValueListCopy);
                Assert.IsTrue(cellValueList.SequenceEqual(cellValueListCopy), "Category column is not sorted in descending order.");
                Results.WriteStatus(test, "Pass", "Verified, Category Column is sorted in descending order.");
            }
            else
            {
                Array.Sort(cellValueListCopy);
                Array.Reverse(cellValueListCopy);
                Assert.IsTrue(cellValueList.SequenceEqual(cellValueListCopy), "Category Column is not sorted in Descending order.");
                Results.WriteStatus(test, "Pass", "Verified, Category Column is sorted in Descending order.");

                driver._click("xpath", "//div[@colid='14']//span[@id='agText']");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@colid='14']//span[contains(@id, 'Sort')]", "id").Contains("Asc"), "Category column is not sorted in Ascending order.");
                driver._scrollintoViewElement("xpath", "//div[@row=0]//div[@colid='14']");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]"), "Category Column cells not present.");
                cellCollection = driver._findElements("xpath", "//div[contains(@class, 'ag-body-viewport')]//div[@colid=14]");

                cellValueList = new string[cellCollection.Count];
                for (int i = 0; i < cellCollection.Count; i++)
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cellCollection[i]);
                    cellValueList[i] = cellCollection[i].Text;
                }

                cellValueListCopy = new string[cellValueList.Length];
                Array.Copy(cellValueList, cellValueListCopy, cellValueList.Length);
                Array.Sort(cellValueListCopy);
                Assert.IsTrue(cellValueList.SequenceEqual(cellValueListCopy), "Category column is not sorted in Ascending order.");
                Results.WriteStatus(test, "Pass", "Verified, Category Column is sorted in Ascending order.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Sort Fucntionality On Detail Data Section");
            return new ManufacturerComparison(driver, test);
        }


        #endregion


    }   
}
