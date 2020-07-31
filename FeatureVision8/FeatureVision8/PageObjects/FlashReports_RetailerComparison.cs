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
    public class FlashReports_RetailerComparison
    {
        #region Private Variables

        private IWebDriver flashReports_RetailerComparison;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public FlashReports_RetailerComparison(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.flashReports_RetailerComparison = driver;
            test = testReturn;
            homePage = new Home(driver, test);
        }

        public IWebDriver driver
        {
            get { return this.flashReports_RetailerComparison; }
            set { this.flashReports_RetailerComparison = value; }
        }

        ///<summary>
        ///Verify Retailer Comparison Screen
        ///</summary>
        ///<returns></returns>
        public FlashReports_RetailerComparison VerifyRetailerComparisonScreen()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='flashreport row']//span[contains(@class, 'page-header')]"), "'Retailer Comparison' header not present.");
            Assert.AreEqual("Retailer Comparison", driver._getText("xpath", "//div[@class='flashreport row']//span[contains(@class, 'page-header')]"), "'Retailer Comparison' header text does not match.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'divExportOption')]"), "Export Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'fa fa-home')]"), "Home Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'all-ads-div')]/span[text()]"), "'Retailer - Comparison' Title not present.");
            Assert.AreEqual("Retailer Comparison", driver._getText("xpath", "//div[contains(@class, 'all-ads-div')]/span[text()]"), "'Retailer - Comparison' Title text does not match.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvFilterField1002']/div[contains(@class, 'filterHeader')]"), "Print Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvFilterField1001']/div[contains(@class, 'filterHeader')]"), "'Country' Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[@title='Open Calendar']"), "'Calendar' icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@type='text' and @id='txtCalendar']"), "Calendar Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'tblRetailer divLeftContainer')]"), "'My Retailer' section not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'divRightContainer')]"), "'Competitor' section not present.");

            Results.WriteStatus(test, "Pass", "Verified, Retailer Retailer Comparison");
            return new FlashReports_RetailerComparison(driver, test);
        }

        ///<summary>
        ///Verify Media/Country DDL
        ///</summary>
        ///<param name="optionName">Option to be selected from DDL</param>
        ///<param name="filterName">Filter to be applied</param>
        ///<returns></returns>
        public FlashReports_RetailerComparison VerifyMedia_CountryDDL(string filterName = "Media", string optionName = "Print")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvFilterField1002']/div[contains(@class, 'filterHeader')]"), "Media Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvFilterField1001']/div[contains(@class, 'filterHeader')]"), "'Country' Field not present.");
            driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

            if (filterName.ToLower().Contains("media"))
            {
                driver._click("xpath", "//div[@id='dvFilterField1002']/div[contains(@class, 'filterHeader')]");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterField1002']//li"), "Media DDL not present.");
                IList<IWebElement> mediaDDLColl = driver._findElements("xpath", "//div[@id='dvFilterField1002']//li");

                bool avail = false;
                foreach (IWebElement mediaDDL in mediaDDLColl)
                    if (mediaDDL.Text.ToLower().Contains(optionName.ToLower()))
                    {
                        avail = true;
                        mediaDDL.Click();
                        break;
                    }
                Assert.IsTrue(avail, "'" + optionName + "' not found.");
                Thread.Sleep(1000);
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterField1003']/div"), "My Retailer Dropdown button not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterField1004']/div"), "Competitor Dropdown button not present.");

                driver._click("xpath", "//div[@id='dvFilterField1003']/div");
                
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterTree1003']//li/div"), "My Retailer DDL not present.");
                IList<IWebElement> ddlCollection = driver._findElements("xpath", "//div[@id='dvFilterTree1003']//li/div");

                avail = true;
                foreach (IWebElement ddlEle in ddlCollection)
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", ddlEle);
                    if (!ddlEle.Text.ToLower().Contains(optionName.ToLower()))
                    {
                        avail = false;
                        break;
                    }
                }

                if (avail)
                {
                    driver._click("xpath", "//div[@id='dvFilterField1004']/div");

                    Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterTree1004']//li/div"), "My Retailer DDL not present.");
                    ddlCollection = driver._findElements("xpath", "//div[@id='dvFilterTree1004']//li/div");

                    avail = true;
                    foreach (IWebElement ddlEle in ddlCollection)
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", ddlEle);
                        if (!ddlEle.Text.ToLower().Contains(optionName.ToLower()))
                        {
                            avail = false;
                            break;
                        }
                    }
                }

                if (optionName.ToLower().Contains("print"))
                {
                    Assert.IsFalse(avail, "'" + optionName + "' Media Filter not applied successfully.");
                    Results.WriteStatus(test, "Pass", "'" + optionName + "' Media Filter applied successfully.");
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
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

                driver._click("xpath", "//div[@id='dvFilterField1003']/div");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterTree1003']//li/div"), "My Retailer DDL not present.");
                IList<IWebElement> ddlCollection = driver._findElements("xpath", "//div[@id='dvFilterTree1003']//li/div");

                avail = true;
                foreach (IWebElement ddlEle in ddlCollection)
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", ddlEle);
                    if (ddlEle.Text.ToLower().Contains("chicago") 
                        || ddlEle.Text.ToLower().Contains("cleveland") || ddlEle.Text.ToLower().Contains("los angeles"))
                    {
                        avail = false;
                        break;
                    }
                }

                if (avail)
                {
                    driver._click("xpath", "//div[@id='dvFilterField1004']/div");

                    Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterTree1003']//li/div"), "My Retailer DDL not present.");
                    ddlCollection = driver._findElements("xpath", "//div[@id='dvFilterTree1004']//li/div");

                    avail = true;
                    foreach (IWebElement ddlEle in ddlCollection)
                    {
                        if (ddlEle.Text.ToLower().Contains("chicago")
                            || ddlEle.Text.ToLower().Contains("cleveland") || ddlEle.Text.ToLower().Contains("los angeles"))
                        {
                            avail = false;
                            break;
                        }
                    }
                }

                if (optionName.ToLower().Contains("canada"))
                {
                    Assert.IsTrue(avail, "'" + optionName + "' Country Filter not applied successfully.");
                    Results.WriteStatus(test, "Pass", "'" + optionName + "' Country Filter applied successfully.");
                }
                else
                {
                    Assert.IsFalse(avail, "'" + optionName + "' Country Filter not applied successfully.");
                    Results.WriteStatus(test, "Pass", "'" + optionName + "' Country Filter applied successfully.");
                }
            }

            Results.WriteStatus(test, "Pass", "Verify '" + filterName + "' DDL");
            return new FlashReports_RetailerComparison(driver, test);
        }

        ///<summary>
        ///Verify Calendar Textbox & Icon
        ///</summary>
        ///<param name="date">Date to be applied</param>
        ///<returns></returns>
        public FlashReports_RetailerComparison VerifyCalendarTextboxAndIcon(string date = "Latest")
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
            if (!driver._getText("xpath", "//div[@id='ui-datepicker-div']//div[2]/div[@class='js_selectedText']").Equals(year))
            {
                driver._click("xpath", "//div[@id='ui-datepicker-div']//div[2]/div[@class='js_selectedText']");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='ui-datepicker-div']//div[2]//li"), "Year DDL not present");
                IList<IWebElement> yearDDLColl = driver._findElements("xpath", "//div[@id='ui-datepicker-div']//div[2]//li");

                bool avail = false;
                foreach (IWebElement yearDDL in yearDDLColl)
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
            DateTime[] prevWeek = new DateTime[7];
            string[] sPrevWeek = new string[7];

            for(int i = 0, j = 0;  j < prevWeek.Length; i--, j++)
            {
                prevWeek[j] = dDate.AddDays(i);
                sPrevWeek[j] = prevWeek[j].ToString("MM/dd/yyyy");
                if (sPrevWeek[j].Contains("-"))
                    sPrevWeek[j] = sPrevWeek[j].Replace("-", "/");
            }

            driver._click("xpath", "//div[@id='dvFilterField1003']/div");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterTree1003']//li/div"), "My Retailer DDL not present.");
            IList<IWebElement> ddlCollection = driver._findElements("xpath", "//div[@id='dvFilterTree1003']//li/div");

            foreach (IWebElement ddlEle in ddlCollection)
            {
                bool avail = false;
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", ddlEle);
                foreach (string prev in sPrevWeek)
                {
                    if (ddlEle.Text.ToLower().Contains(prev))
                    {
                        avail = true;
                        break;
                    }
                }
                Assert.IsTrue(avail, "Calendar Week not applied.");
            }

            driver._click("xpath", "//div[@id='dvFilterField1004']/div");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvFilterTree1004']//li/div"), "My Retailer DDL not present.");
            ddlCollection = driver._findElements("xpath", "//div[@id='dvFilterTree1004']//li/div");

            foreach (IWebElement ddlEle in ddlCollection)
            {
                bool avail = false;
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", ddlEle);
                foreach (string prev in sPrevWeek)
                {
                    if (ddlEle.Text.ToLower().Contains(prev))
                    {
                        avail = true;
                        break;
                    }
                }
                Assert.IsTrue(avail, "Calendar Week not applied.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Calendar Textbox & Icon for date applied calendar week.");
            return new FlashReports_RetailerComparison(driver, test);
        }

        ///<summary>
        ///Verify Export Options
        ///</summary>
        ///<param name="optionName">Option to be selected</param>
        ///<returns></returns>
        public FlashReports_RetailerComparison VerifyExportOptions(string optionName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'divExportOption')]"), "Export Icon not present.");

            if (optionName.ToLower().Contains("png"))
            {
                driver._click("xpath", "//div[contains(@class,'divExportOption')]");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[@id='ulExportOption']/li[text()='Export as PNG']"), "Export DDL option 'Export as PNG' not present.");
                driver.MouseHoverUsingElement("xpath", "//ul[@id='ulExportOption']/li[text()='Export as PNG']");
                driver._clickByJavaScriptExecutor("//ul[@id='ulExportOption']/li[text()='Export as PNG']");
            }
            else if (optionName.ToLower().Contains("jpg"))
            {
                driver._click("xpath", "//div[contains(@class,'divExportOption')]");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[@id='ulExportOption']/li[text()='Export as JPG']"), "Export DDL option 'Export as JPG' not present.");
                driver.MouseHoverUsingElement("xpath", "//ul[@id='ulExportOption']/li[text()='Export as JPG']");
                driver._clickByJavaScriptExecutor("//ul[@id='ulExportOption']/li[text()='Export as JPG']");
            }
            else if (optionName.ToLower().Contains("pdf"))
            {
                driver._click("xpath", "//div[contains(@class,'divExportOption')]");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[@id='ulExportOption']/li[text()='Export as PDF']"), "Export DDL option 'Export as PDF' not present.");
                driver.MouseHoverUsingElement("xpath", "//ul[@id='ulExportOption']/li[text()='Export as JPG']");
                driver._clickByJavaScriptExecutor("//ul[@id='ulExportOption']/li[text()='Export as PDF']");
            }
            Thread.Sleep(5000);

            driver._waitForElementToBeHidden("xpath", "//div[@id='dvContent1']/div", 600);

            Results.WriteStatus(test, "Pass", "Verified, '" + optionName + "' Export Option");
            return new FlashReports_RetailerComparison(driver, test);
        }

        ///<summary>
        ///Verify My Retailer Section
        ///</summary>
        ///<param name="retailer">Retailer or Competitor Section</param>
        ///<returns></returns>
        public FlashReports_RetailerComparison VerifyMyRetailerSection(bool retailer = true)
        {
            string classPart = "divRight", sectionName = "Competitor";
            if (retailer)
            {
                classPart = "divLeft";
                sectionName = "My Retailer";
            }
            //Assert.IsTrue(driver._waitForElement("xpath", ""), "'" + sectionName + " Header' not present.");
            //Assert.AreEqual(sectionName.ToLower(), driver._getText("xpath", "").ToLower(), "'My Retailer Header' text does not match.");


            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='divRetCompHeaderDetail']/span"), "'" + sectionName + "' DDL not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//img"), "Page Image not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//span[@title='View Page Full Size']"), "'View Page Full Size' text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//span[contains(@class,'ViewPageFullSizeIcon')]"), "'View Page Full Size' icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']"), "'Pagination' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//table"), "'Ad Info' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//table//td[1]"), "'Ad Info' type not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//table//td[2]"), "'Ad Info' value present.");

            Results.WriteStatus(test, "Pass", "Verified, " + sectionName + " Section");
            return new FlashReports_RetailerComparison(driver, test);
        }

        ///<summary>
        ///Verify Retailer Dropdown
        ///</summary>
        ///<param name="retailer">Retailer or Competitor Section</param>
        ///<returns></returns>
        public FlashReports_RetailerComparison VerifyRetailerDropdown(bool retailer = true)
        {
            string classPart = "divRight", sectionName = "Competitor";
            if (retailer)
            {
                classPart = "divLeft";
                sectionName = "My Retailer";
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='divRetCompHeaderDetail']/span"), "'" + sectionName + "' DDL not present.");
            string currRetailer = driver._getText("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='divRetCompHeaderDetail']/span");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, '" + classPart + "')]//div[contains(@id,'dvFilterField100')]/div"), "My Retailer Dropdown button not present.");

            driver._click("xpath", "//div[contains(@class, '" + classPart + "')]//div[contains(@id,'dvFilterField100')]/div");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, '" + classPart + "')]//div[contains(@id,'dvFilterTree100')]//li/div"), "My Retailer DDL not present.");
            IList<IWebElement> ddlCollection = driver._findElements("xpath", "//div[contains(@class, '" + classPart + "')]//div[contains(@id,'dvFilterTree100')]//li/div");

            Random rand = new Random();
            int x = rand.Next(0, ddlCollection.Count);
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", ddlCollection[x]);
            string newRetailer = ddlCollection[x].Text;
            ddlCollection[x].Click();
            Thread.Sleep(2000);
            driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='divRetCompHeaderDetail']/span"), "'" + sectionName + "' DDL not present.");
            Assert.IsFalse(currRetailer.Contains(driver._getText("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='divRetCompHeaderDetail']/span")), "Retailer is not changed successfully.");
            Assert.IsTrue(newRetailer.Contains(driver._getText("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='divRetCompHeaderDetail']/span")), "Retailer is not changed successfully.");

            Results.WriteStatus(test, "Pass", "Verified, Retailer Dropdown of '" + sectionName + "' section");
            return new FlashReports_RetailerComparison(driver, test);
        }

        ///<summary>
        ///Verify Pagination
        ///</summary>
        ///<param name="input">Page to navigate to</param>
        ///<param name="retailer">Retailer or Competitor Section</param>
        ///<returns></returns>
        public FlashReports_RetailerComparison VerifyPagination(string input = "", bool retailer = true)
        {
            string classPart = "divRight", sectionName = "Competitor";
            if (retailer)
            {
                classPart = "divLeft";
                sectionName = "My Retailer";
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']"), "'Pagination' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//span[@title='NEXT PAGE']"), "Next button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//span[text()='NEXT PAGE']"), "Next Page Text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//span[@title='PREV PAGE']"), "Prev button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//span[text()='PREV PAGE']"), "Prev Page Text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input"), "Page No. Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//label"), "Maximum No. of Pages Text not present.");
            int iPrevPage = 0; int iNextPage = 0;int iMaxPage = 0;int intInput = 0;int iNewPage = 0;

            if (input.ToLower().Contains("prev"))
            {
                Assert.IsFalse(driver._getAttributeValue("xpath", "//div[contains(@class, '" + classPart + "')]//div[contains(@class,'FullSizeImagePrevPage')]", "class").Contains("disable"), "Previous button is disabled.");
                string prevPage = driver._getValue("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input");
                Assert.IsTrue(int.TryParse(prevPage, out iPrevPage), "Couldn't convert '" + prevPage + "' to int.");
                driver._click("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//span[@title='PREV PAGE']");
                Thread.Sleep(2000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input"), "Page No. Field not present.");
                string nextPage = driver._getValue("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input");
                Assert.IsTrue(int.TryParse(nextPage, out iNextPage), "Couldn't convert '" + nextPage + "' to int.");
                Assert.AreEqual(iNextPage + 1, iPrevPage, "Prev Button did not work.");
            }
            else if (input.ToLower().Contains("next"))
            {
                Assert.IsFalse(driver._getAttributeValue("xpath", "//div[contains(@class, '" + classPart + "')]//div[contains(@class,'FullSizeImageNextPage')]", "class").Contains("disable"), "Previous button is disabled.");
                string prevPage = driver._getValue("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input");
                Assert.IsTrue(int.TryParse(prevPage, out iPrevPage), "Couldn't convert '" + prevPage + "' to int.");
                driver._click("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//span[@title='NEXT PAGE']");
                Thread.Sleep(2000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input"), "Page No. Field not present.");
                string nextPage = driver._getValue("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input");
                Assert.IsTrue(int.TryParse(nextPage, out iNextPage), "Couldn't convert '" + nextPage + "' to int.");
                Assert.AreEqual(iNextPage, iPrevPage + 1, "Next Button did not work.");
            }
            else
            {
                string maxPage = driver._getText("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//label");
                Assert.IsTrue(int.TryParse(maxPage, out iMaxPage), "Couldn't convert '" + maxPage + "' to int.");
                Assert.IsTrue(int.TryParse(input, out intInput), "Couldn't convert '" + input + "' to int.");
                Assert.LessOrEqual(intInput, iMaxPage, "Input Page No. is greater than maximum page.");
                driver._click("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input");
                driver._clearText("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input");
                driver._type("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input", input);

                driver._click("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//label");
                Thread.Sleep(2000);
                string newPage = driver._getValue("xpath", "//div[contains(@class, '" + classPart + "')]//div[@class='FullSizeImagePagerDiv']//input");
                Assert.IsTrue(int.TryParse(newPage, out iNewPage), "Couldn't convert '" + newPage + "' to int.");
                Assert.AreEqual(intInput, iNewPage, "Input did not work.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Pagination on '" + sectionName + "' section.");
            return new FlashReports_RetailerComparison(driver, test);
        }

        ///<summary>
        ///Verify View Page Full Size Link
        ///</summary>
        ///<param name="retailer">Retailer or Competitor Section</param>
        ///<returns></returns>
        public FlashReports_RetailerComparison VerifyViewPageFullSizeLink(bool retailer = true)
        {
            string classPart = "divRight", sectionName = "Competitor";
            if (retailer)
            {
                classPart = "divLeft";
                sectionName = "My Retailer";
            }

            string retailerName = "";
            string retailerPart = "";

            if (!retailer)
            {
                retailerName = driver._getText("xpath", "//div[contains(@class, '" + classPart + "')]//div[contains(@class, 'competitorHeader')]");
                retailerPart = driver._getText("xpath", "//div[contains(@class, '" + classPart + "')]//span[contains(@class, 'competitorDetail')]");
            }
            else
            {
                retailerName = driver._getText("xpath", "//div[contains(@class, '" + classPart + "')]//div[contains(@class, 'retailerHeader')]");
                retailerPart = driver._getText("xpath", "//div[contains(@class, '" + classPart + "')]//span[contains(@class, 'retailerDetail')]");
            }
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//span[@title='View Page Full Size']"), "'View Page Full Size' text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, '" + classPart + "')]//span[contains(@class,'ViewPageFullSizeIcon')]"), "'View Page Full Size' icon not present.");

            driver._click("xpath", "//div[contains(@class, '" + classPart + "')]//span[@title='View Page Full Size']");
            Thread.Sleep(2000);
            driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='flashreport row']//span[contains(@class, 'page-header')]"), "'Retailer Comparison' header not present.");
            Assert.AreEqual("Retailer Comparison", driver._getText("xpath", "//div[@class='flashreport row']//span[contains(@class, 'page-header')]"), "'Retailer Comparison' header text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'fa fa-home')]"), "Home Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'all-ads-div')]/span[text()]"), "'back To Retailer Comparison' Title not present.");
            Assert.AreEqual("Back To Retailer Comparison", driver._getText("xpath", "//div[contains(@class, 'all-ads-div')]/span[text()]"), "'Back To Retailer Comparison' Title text does not match.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'AllPagestext')]"), "'Retailer - Name' Title not present in title.");
            Assert.IsTrue(driver._getText("xpath", "//span[contains(@class, 'AllPagestext')]").Contains(retailerName)
                && driver._getText("xpath", "//span[contains(@class, 'AllPagestext')]").Contains(retailerPart), "'Retailer Name' text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class= 'FullPageViewHeaderText']"), "'FULL PAGE VIEW' Title not present.");
            Assert.AreEqual("FULL PAGE VIEW", driver._getText("xpath", "//div[@class= 'FullPageViewHeaderText']"), "'FULL PAGE VIEW' Title text does not match.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class= 'FullPageViewHeaderReturn']"), "'RETURN TO NORMAL VIEW' Title not present.");
            Assert.IsTrue(driver._getText("xpath", "//div[@class= 'FullPageViewHeaderReturn']").Contains("RETURN TO NORMAL VIEW"), "'RETURN TO NORMAL VIEW' Title text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='FullPageViewMainPager']//span[@title='NEXT PAGE']"), "Next Page icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='FullPageViewMainPager']//span[text()='NEXT PAGE']"), "Next Page Text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='FullPageViewMainPager']//span[@title='PREV PAGE']"), "Prev Page icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='FullPageViewMainPager']//span[text()='PREV PAGE']"), "Prev Page Text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='FullSizeImagePager']//input"), "Page No. Field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='FullSizeImagePager']//label"), "Maximum Page No. Label not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'FullPageViewMainImage')]//img"), "Full Page Image not present.");

            Assert.AreEqual("1", driver._getValue("xpath", "//div[@class='FullSizeImagePager']//input"), "Page 1 is not loaded in Full Page View.");
            driver._click("xpath", "//div[@class='FullPageViewMainPager']//span[@title='NEXT PAGE']");
            Thread.Sleep(3000);
            Assert.AreEqual("2", driver._getValue("xpath", "//div[@class='FullSizeImagePager']//input"), "Page 1 is not loaded in Full Page View.");
            driver._click("xpath", "//div[@class='FullPageViewMainPager']//span[@title='PREV PAGE']");
            Thread.Sleep(3000);
            Assert.AreEqual("1", driver._getValue("xpath", "//div[@class='FullSizeImagePager']//input"), "Page 1 is not loaded in Full Page View.");

            driver._click("xpath", "//div[@class= 'FullPageViewHeaderReturn']");

            Results.WriteStatus(test, "Pass", "Verified, View Page Full Size Link on '" + sectionName + "' section.");
            return new FlashReports_RetailerComparison(driver, test);
        }

        #endregion
    }
}
