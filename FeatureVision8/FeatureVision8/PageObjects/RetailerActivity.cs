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
    public class RetailerActivity
    {
        #region Private Variables

        private IWebDriver retailerActivity;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public RetailerActivity(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.retailerActivity = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.retailerActivity; }
            set { this.retailerActivity = value; }
        }

        ///<summary>
        ///Verify Retailer Activity Screen
        ///</summary>
        ///<param name="searchName">Search Name loaded</param>
        ///<returns></returns>
        public RetailerActivity VerifyRetailerActivityScreen(bool verifySections = false, string clientName = "Proctor & Gamble", string searchName = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//h1[text()='Retailer Activity']"), "Retailer Activity Screen not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='title cursorpointer']/a"), "Retailer Circular Strategy link not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-media']/i"), "Retailer Circular Strategy icon not present.");
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
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='2']"), "Madlib Search Parameter 'Any Retailer' not present.");
            if (!clientName.ToLower().Contains("australia"))
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='3']"), "Madlib Search Parameter 'Any Market' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='4']"), "Madlib Search Parameter 'Any Date' not present.");

            Assert.IsTrue(driver._waitForElement("xpath", "//li[@heading]"), "Tabs not present.");
            string[] tabNameList = new string[] { "Ads by Channel", "Pages by Channel", "Ads by Week", "Events by Week" };
            if (clientName.ToLower().Contains("canada"))
                tabNameList[3] = "Themes by Week";
            else if (clientName.ToLower().Contains("australia"))
                Array.Resize(ref tabNameList, 3);
            IList<IWebElement> tabCollection = driver._findElements("xpath", "//li[@heading]");

            foreach (string tabName in tabNameList)
            {
                bool avail = false;
                foreach (IWebElement tab in tabCollection)
                    if (tab.Text.ToLower().Contains(tabName.ToLower()))
                    {
                        avail = true;
                        break;
                    }
                Assert.IsTrue(avail, "'" + tabName + "' tab not found.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Save Options']/i"), "'Search' option not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-sidebar//div[@class='iscsidebar-container']"), "'Navigation' Sidebar not present.");

            if (verifySections)
            {
                if (!driver._getAttributeValue("xpath", "//li[@heading='Ads by Channel']", "class").Contains("active"))
                    driver._click("xpath", "//li[@heading='Ads by Channel']");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//li[@heading='Ads by Channel']", "class").Contains("active"), "'Ads by Channel' tab is not selected.");

                if (!clientName.ToLower().Contains("australia"))
                {
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Channel']"), "Channel radio button not present.");
                    if (clientName.ToLower().Contains("canada"))
                        Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Account Group']"), "Account Group radio button not present.");
                    else
                        Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Parent Retailer']"), "Parent Retailer radio button not present.");
                }

                driver._click("xpath", "//li[@heading='Pages by Channel']");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//li[@heading='Pages by Channel']", "class").Contains("active"), "'Ads by Channel' tab is not selected.");

                if (!clientName.ToLower().Contains("australia"))
                {
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Channel']"), "Channel radio button not present.");
                    if (clientName.ToLower().Contains("canada"))
                        Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Account Group']"), "Account Group radio button not present.");
                    else
                        Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Parent Retailer']"), "Parent Retailer radio button not present.");
                }

                driver._click("xpath", "//li[@heading='Ads by Week']");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//li[@heading='Ads by Week']", "class").Contains("active"), "'Ads by Week' tab is not selected.");

                if (!clientName.ToLower().Contains("australia"))
                {
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Ads']"), "Ads radio button not present.");
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Pages']"), "Pages radio button not present.");
                }

                if (driver._isElementPresent("xpath", "//li[@heading='Themes by Week']"))
                {
                    driver._click("xpath", "//li[@heading='Themes by Week']");
                    Thread.Sleep(1000);
                    Assert.IsTrue(driver._getAttributeValue("xpath", "//li[@heading='Themes by Week']", "class").Contains("active"), "'Themes by Week' tab is not selected.");

                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Event']"), "Event radio button not present.");
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Theme']"), "Theme radio button not present.");
                }
            }

            Results.WriteStatus(test, "Pass", "Verified, Retailer Activity Screen");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Select Tab On Retailer Activity
        ///</summary>
        ///<param name="tabName">Tab to select</param>
        ///<returns></returns>
        public RetailerActivity SelectTabOnRetailerActivity(string tabName = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//li[@heading]"), "Tabs not present.");
            IList<IWebElement> tabCollection = driver._findElements("xpath", "//li[@heading]");

            bool avail = false;
            foreach (IWebElement tab in tabCollection)
                if (tab.Text.ToLower().Contains(tabName.ToLower()))
                {
                    avail = true;
                    tab.Click();
                    break;
                }
            Assert.IsTrue(avail, "'" + tabName + "' tab not found.");

            Thread.Sleep(2000);

            Assert.IsTrue(driver._getAttributeValue("xpath", "//li[@heading='" + tabName + "']", "class").Contains("active"), "'" + tabName + "' tab is not selected.");

            Results.WriteStatus(test, "Pass", "Selected, Tab '" + tabName + "' On Retailer Activity");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Export Menu and Select Option
        ///</summary>
        ///<param name="optionName">Option to be selected</param>
        ///<returns></returns>
        public RetailerActivity VerifyExportMenuAndSelectOption(string optionName = "")
        {
            string chartHeaderId = "";
            if (driver._isElementPresent("xpath", "//div[@id='FlyerColumnChannelCustChartHeader']"))
                chartHeaderId = "FlyerColumnChannelCustChartHeader";
            else if (driver._isElementPresent("xpath", "//div[@id='FlyerColumnTradeCustChartHeader']"))
                chartHeaderId = "FlyerColumnTradeCustChartHeader";
            else if (driver._isElementPresent("xpath", "//div[@id='dvAccountMarketFlyerGridHeader']"))
                chartHeaderId = "dvAccountMarketFlyerGridHeader";
            else if (driver._isElementPresent("xpath", "//div[@id='dvAccountMarketPageGridHeader']"))
                chartHeaderId = "dvAccountMarketPageGridHeader";
            else if (driver._isElementPresent("xpath", "//div[@id='dvAccountMarketEventGridHeader']"))
                chartHeaderId = "dvAccountMarketEventGridHeader";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//div[@title='Export']"), "Export Button not present.");

            string[] exportOptionList = new string[] { "Download PNG", "Download JPG", "Download PDF", "Download EXCEL", "Download PowerPoint" };
            Actions action = new Actions(driver);
            IWebElement exportButton = null;
            driver._click("xpath", "//div[@id='" + chartHeaderId + "']//div[@title='Export']");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='1']//a"), "Export DDL not present.");

            foreach (string exportOption in exportOptionList)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='1']//li/a[text()='" + exportOption + "']"), "'" + exportOption + "' Option not present");
                action.MoveToElement(driver.FindElement(By.XPath("//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='1']//li/a[text()='" + exportOption + "']"))).MoveByOffset(2, 1).Perform();
                if (exportOption.ToLower().Equals(optionName.ToLower()))
                    exportButton = driver.FindElement(By.XPath("//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='1']//li/a[text()='" + exportOption + "']"));
            }

            if (optionName != "")
            {
                Assert.AreNotEqual(null, exportButton, "'" + optionName + "' not found.");
                exportButton.Click();
                Thread.Sleep(10000);
            }

            Results.WriteStatus(test, "Pass", "Verified, Export Menu and Selected Option '" + optionName + "'");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Ads by Channel/Parent Retailer Section
        ///</summary>
        ///<param name="channel">Whether The Section is Ads by Channel</param>
        ///<returns></returns>
        public RetailerActivity VerifyAdsByChannel_ParentRetailerSection(bool channel = true)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Channel']"), "Channel radio button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Parent Retailer']"), "Parent Retailer radio button not present.");
            driver._scrollintoViewElement("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Channel']");
            string chartId = "";

            if (channel)
            {
                driver._click("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Channel']");
                Thread.Sleep(4000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='FlyerColumnChannelCustChartHeaderTitle']//span"), "Section header not present.");
                driver._scrollintoViewElement("xpath", "//div[@id='FlyerColumnChannelCustChartHeaderTitle']//span");
                Assert.AreEqual("Ads by Channel", driver._getText("xpath", "//div[@id='FlyerColumnChannelCustChartHeaderTitle']//span"), "'Ads by Channel' header text does not match.");
                chartId = "FlyerColumnChannelCustChart0chart";
            }
            else
            {
                driver._click("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Parent Retailer']");
                Thread.Sleep(4000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='FlyerColumnTradeCustChartHeaderTitle']//span"), "Section header not present.");
                driver._scrollintoViewElement("xpath", "//div[@id='FlyerColumnTradeCustChartHeaderTitle']//span");
                Assert.AreEqual("Ads by Parent Retailer", driver._getText("xpath", "//div[@id='FlyerColumnTradeCustChartHeaderTitle']//span"), "'Ads by Parent Retailer' header text does not match.");
                chartId = "FlyerColumnTradeCustChart0chart";
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-container']//*[name()='svg']"), "Column Chart not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-yaxis-labels']/span"), "Y Axis not present in Column Chart.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span"), "x Axis not present in Column Chart.");
            if (driver._getText("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-yaxis-labels']/span").Contains("%"))
                Results.WriteStatus(test, "Pass", "Column Chart is displayed as % of Ads");
            else
                Results.WriteStatus(test, "Pass", "Column Chart is displayed as # of Ads");

            if (channel)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='FlyerColumnChannelCustChart0chartTabular_tblMain']"), "Data Table not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='FlyerColumnChannelCustChart0chartTabular_tblMain']//tr[@class]/td"), "Header Cells not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='FlyerColumnChannelCustChart0chartTabular_tblMain']//tr[@class]/td[text()='Ads by Channel']"), "Header Cell for 'Ads by Channel' not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='FlyerColumnChannelCustChart0chartTabular_tblMain']//tr[@class]/td[text()='Values']"), "Header Cell for 'Values' not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='FlyerColumnChannelCustChart0chartTabular_tblMain']//tr[not(@class)]/td"), "Data Cells not present in Data Table");
            }
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='FlyerColumnTradeCustChart0chartTabular_FixColumnHeader']"), "Data Table not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='FlyerColumnTradeCustChart0chartTabular_FixColumnHeader']//tr[@class]/td"), "Header Cells not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='FlyerColumnTradeCustChart0chartTabular_FixColumnHeader']//tr[@class]/td[text()='Ads by Parent Retailer']"), "Header Cell for 'Ads by Channel' not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='FlyerColumnTradeCustChart0chartTabular_FixMainHeader']//tr[@class]/td[text()='Values']"), "Header Cell for 'Values' not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[contains(@id, 'FlyerColumnTradeCustChart0')]//tr[not(@class)]/td"), "Data Cells not present in Data Table");
            }


            Results.WriteStatus(test, "Pass", "Verified, Ads by Channel/Parent Retailer Section.");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Help Icon on Ads By Channel or Parent Retailer Section
        ///</summary>
        ///<returns></returns>
        public RetailerActivity VerifyHelpIconOnAdsByChannelOrParentRetailerSection()
        {
            string chartHeaderId = "";
            if (driver._isElementPresent("xpath", "//div[@id='FlyerColumnChannelCustChartHeader']"))
                chartHeaderId = "FlyerColumnChannelCustChartHeader";
            else
                chartHeaderId = "FlyerColumnTradeCustChartHeader";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//div[@class='fa fa-question homeChartSearch']"), "Help Button not present.");
            driver.MouseHoverUsingElement("xpath", "//div[@id='" + chartHeaderId + "']//div[@class='fa fa-question homeChartSearch']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='" + chartHeaderId + "']//div[@onmouseout='HideTooltip();']"), "Help Icon Tooltip not present.");
            string tooltipText = "Who increased or decreased circular drops (indicating a shift in strategy)?</b><br/><br/>Provides a total by channel values (Drug, Mass, Food) in a given period, compared to last period (if selected). Ads are defined as: Circular, FSI, Mailer, Newspaper Media Types only. If ‘Product’ level selections are made, it will ignore these selections and return ad level results, i.e. total number of ads/pages in a Circular, regardless of categories advertised.";
            Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@id='" + chartHeaderId + "']//div[@onmouseout='HideTooltip();']", "onmouseover").Contains(tooltipText), "Help Icon Text does not match.");

            Results.WriteStatus(test, "Pass", "Verified, Help Icon on Ads By Channel or Parent Retailer Section");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Window Icon on Ads By Channel or Parent Retailer Section
        ///</summary>
        ///<param name="reset">Whether Reset should be disabled</param>
        ///<param name="selectOption">Option to be selected</param>
        ///<returns></returns>
        public RetailerActivity VerifyWindowIconOnAdsByChannelOrParentRetailerSection(bool reset = false, string selectOption = "")
        {
            string chartHeaderId = "";
            if (driver._isElementPresent("xpath", "//div[@id='FlyerColumnChannelCustChartHeader']"))
                chartHeaderId = "FlyerColumnChannelCustChartHeader";
            else
                chartHeaderId = "FlyerColumnTradeCustChartHeader";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//div[@title='More Options']"), "Window Icon not present.");
            string[] windowOptionList = new string[] { "Show by #", "Show by %", "Reset" };
            Actions action = new Actions(driver);
            IWebElement windowButton = null;
            driver._click("xpath", "//div[@id='" + chartHeaderId + "']//div[@title='More Options']");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//a"), "Export DDL not present.");

            foreach (string windowOption in windowOptionList)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//li/a[text()='" + windowOption + "']"), "'" + windowOption + "' Option not present");
                if (reset && windowOption.ToLower().Equals("reset"))
                    Assert.IsFalse(driver._getAttributeValue("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//li/a[text()='" + windowOption + "']", "class").Contains("disable"), "Reset button is not disabled");
                action.MoveToElement(driver.FindElement(By.XPath("//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//li/a[text()='" + windowOption + "']"))).MoveByOffset(2, 1).Perform();
                if (windowOption.ToLower().Equals(selectOption.ToLower()))
                    windowButton = driver.FindElement(By.XPath("//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//li/a[text()='" + windowOption + "']"));
            }

            if (selectOption != "")
            {
                Assert.AreNotEqual(null, windowButton, "'" + selectOption + "' not found.");
                windowButton.Click();
                Thread.Sleep(10000);
            }

            Results.WriteStatus(test, "Pass", "Verified, Help Icon on Ads By Channel or Parent Retailer Section");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify When Column Is Drilled Down On
        ///</summary>
        ///<returns></returns>
        public RetailerActivity VerifyWhenColumnIsDrilledDownOn()
        {
            string chartId = "";
            if (driver._isElementPresent("xpath", "//div[@id='FlyerColumnChannelCustChartHeader']"))
                chartId = "FlyerColumnChannelCustChart0chart";
            else
                chartId = "FlyerColumnTradeCustChart0chart";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span"), "x Axis not present in Column Chart.");
            IList<IWebElement> xAxisColl = driver._findElements("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span");
            string columnName = xAxisColl[0].Text;
            columnName = columnName.TrimStart(' ');

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//*[name()='rect' and not(@x=0) and not(@y=0)]"), "Bar Graphs not present.");
            IList<IWebElement> barCollection = driver._findElements("xpath", "//div[@id='" + chartId + "']//*[name()='rect' and not(@x=0) and not(@y=0)]");
            barCollection[0].Click();
            Thread.Sleep(2000);

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='FlyerColumnChannelCustChartHeader']//li[text()]"), "New Header not present.");
            Assert.AreEqual(columnName.ToLower() + " - retailer", driver._getText("xpath", "//div[@id='FlyerColumnChannelCustChartHeader']//li[text()]").ToLower(), "Column Name is not updated in the header.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span"), "x Axis not present in Column Chart.");
            xAxisColl = driver._findElements("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span");
            Assert.AreNotEqual(columnName, xAxisColl[0].Text, "Column Chart not updated.");

            Results.WriteStatus(test, "Pass", "Verified, When Column Is Drilled Down On");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Export Menu and Select Option On Pages by Channel/Parent Retailer
        ///</summary>
        ///<returns></returns>
        public RetailerActivity VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer(string optionName = "")
        {
            string chartHeaderId = "";
            if (driver._isElementPresent("xpath", "//div[@id='PageColumnChannelCustChartHeader']"))
                chartHeaderId = "PageColumnChannelCustChartHeader";
            else
                chartHeaderId = "PageColumnTradeCustChartHeader";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//div[@title='Export']"), "Export Button not present.");

            string[] exportOptionList = new string[] { "Download PNG", "Download JPG", "Download PDF", "Download EXCEL", "Download PowerPoint" };
            Actions action = new Actions(driver);
            IWebElement exportButton = null;
            driver._click("xpath", "//div[@id='" + chartHeaderId + "']//div[@title='Export']");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='1']//a"), "Export DDL not present.");

            foreach (string exportOption in exportOptionList)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='1']//li/a[text()='" + exportOption + "']"), "'" + exportOption + "' Option not present");
                action.MoveToElement(driver.FindElement(By.XPath("//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='1']//li/a[text()='" + exportOption + "']"))).MoveByOffset(2, 1).Perform();
                if (exportOption.ToLower().Equals(optionName.ToLower()))
                    exportButton = driver.FindElement(By.XPath("//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='1']//li/a[text()='" + exportOption + "']"));
            }

            if (optionName != "")
            {
                Assert.AreNotEqual(null, exportButton, "'" + optionName + "' not found.");
                exportButton.Click();
                Thread.Sleep(10000);
            }

            Results.WriteStatus(test, "Pass", "Verified, Export Menu and Selected Option '" + optionName + "' On Pages by Channel/Parent Retailer");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Pages by Channel/Parent Retailer Section
        ///</summary>
        ///<param name="channel">Whether The Section is Pages by Channel</param>
        ///<returns></returns>
        public RetailerActivity VerifyPagesByChannel_ParentRetailerSection(bool channel = true)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Channel']"), "Channel radio button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Parent Retailer']"), "Parent Retailer radio button not present.");
            driver._scrollintoViewElement("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Channel']");
            string chartId = "";

            if (channel)
            {
                driver._click("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Channel']");
                Thread.Sleep(4000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='PageColumnChannelCustChartHeaderTitle']//span"), "Section header not present.");
                driver._scrollintoViewElement("xpath", "//div[@id='PageColumnChannelCustChartHeaderTitle']//span");
                Assert.AreEqual("Pages by Channel", driver._getText("xpath", "//div[@id='PageColumnChannelCustChartHeaderTitle']//span"), "'Pages by Channel' header text does not match.");
                chartId = "PageColumnChannelCustChart0chart";
            }
            else
            {
                driver._click("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Parent Retailer']");
                Thread.Sleep(4000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='PageColumnTradeCustChartHeaderTitle']//span"), "Section header not present.");
                driver._scrollintoViewElement("xpath", "//div[@id='PageColumnTradeCustChartHeaderTitle']//span");
                Assert.AreEqual("Pages by Parent Retailer", driver._getText("xpath", "//div[@id='PageColumnTradeCustChartHeaderTitle']//span"), "'Pages by Parent Retailer' header text does not match.");
                chartId = "PageColumnTradeCustChart0chart";
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-container']//*[name()='svg']"), "Column Chart not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-yaxis-labels']/span"), "Y Axis not present in Column Chart.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span"), "x Axis not present in Column Chart.");
            if (driver._getText("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-yaxis-labels']/span").Contains("%"))
                Results.WriteStatus(test, "Pass", "Column Chart is displayed as % of Pages");
            else
                Results.WriteStatus(test, "Pass", "Column Chart is displayed as # of Pages");

            if (channel)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='PageColumnChannelCustChart0chartTabular_tblMain']"), "Data Table not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='PageColumnChannelCustChart0chartTabular_tblMain']//tr[@class]/td"), "Header Cells not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='PageColumnChannelCustChart0chartTabular_tblMain']//tr[@class]/td[text()='Pages by Channel']"), "Header Cell for 'Pages by Channel' not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='PageColumnChannelCustChart0chartTabular_tblMain']//tr[@class]/td[text()='Values']"), "Header Cell for 'Values' not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='PageColumnChannelCustChart0chartTabular_tblMain']//tr[not(@class)]/td"), "Data Cells not present in Data Table");
            }
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='PageColumnTradeCustChart0chartTabular_FixColumnHeader']"), "Data Table not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='PageColumnTradeCustChart0chartTabular_FixColumnHeader']//tr[@class]/td"), "Header Cells not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='PageColumnTradeCustChart0chartTabular_FixColumnHeader']//tr[@class]/td[text()='Pages by Parent Retailer']"), "Header Cell for 'Pages by Channel' not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='PageColumnTradeCustChart0chartTabular_FixMainHeader']//tr[@class]/td[text()='Values']"), "Header Cell for 'Values' not present in Data Table");
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[contains(@id, 'PageColumnTradeCustChart0')]//tr[not(@class)]/td"), "Data Cells not present in Data Table");
            }


            Results.WriteStatus(test, "Pass", "Verified, Pages by Channel/Parent Retailer Section.");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Help Icon on Pages By Channel or Parent Retailer Section
        ///</summary>
        ///<returns></returns>
        public RetailerActivity VerifyHelpIconOnPagesByChannelOrParentRetailerSection()
        {
            string chartHeaderId = "";
            if (driver._isElementPresent("xpath", "//div[@id='PageColumnChannelCustChartHeader']"))
                chartHeaderId = "PageColumnChannelCustChartHeader";
            else
                chartHeaderId = "PageColumnTradeCustChartHeader";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//div[@class='fa fa-question homeChartSearch']"), "Help Button not present.");
            driver.MouseHoverUsingElement("xpath", "//div[@id='" + chartHeaderId + "']//div[@class='fa fa-question homeChartSearch']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='" + chartHeaderId + "']//div[@onmouseout='HideTooltip();']"), "Help Icon Tooltip not present.");
            string tooltipText = "<b>Who increased or decreased circular page counts (indicating a shift in strategy)?</b><br/><br/>Provides a total by channel values (Drug, Mass, Food) in a given period, compared to last period (if selected). Ads are defined as: Circular, FSI, Mailer, Newspaper Media Types only. If ‘Product’ level selections are made, it will ignore these selections and return ad level results, i.e. total number of ads/pages in a Circular, regardless of categories advertised.";
            Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@id='" + chartHeaderId + "']//div[@onmouseout='HideTooltip();']", "onmouseover").Contains(tooltipText), "Help Icon Text does not match.");

            Results.WriteStatus(test, "Pass", "Verified, Help Icon on Pages By Channel or Parent Retailer Section");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Window Icon on Pages By Channel or Parent Retailer Section
        ///</summary>
        ///<param name="reset">Whether Reset should be disabled</param>
        ///<param name="selectOption">Option to be selected</param>
        ///<returns></returns>
        public RetailerActivity VerifyWindowIconOnPagesByChannelOrParentRetailerSection(bool reset = false, string selectOption = "")
        {
            string chartHeaderId = "";
            if (driver._isElementPresent("xpath", "//div[@id='PageColumnChannelCustChartHeader']"))
                chartHeaderId = "PageColumnChannelCustChartHeader";
            else
                chartHeaderId = "PageColumnTradeCustChartHeader";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//div[@title='More Options']"), "Window Icon not present.");
            string[] windowOptionList = new string[] { "Show by #", "Show by %", "Reset" };
            Actions action = new Actions(driver);
            IWebElement windowButton = null;
            driver._click("xpath", "//div[@id='" + chartHeaderId + "']//div[@title='More Options']");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//a"), "Export DDL not present.");

            foreach (string windowOption in windowOptionList)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//li/a[text()='" + windowOption + "']"), "'" + windowOption + "' Option not present");
                if (reset && windowOption.ToLower().Equals("reset"))
                    Assert.IsFalse(driver._getAttributeValue("xpath", "//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//li/a[text()='" + windowOption + "']", "class").Contains("disable"), "Reset button is not disabled");
                action.MoveToElement(driver.FindElement(By.XPath("//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//li/a[text()='" + windowOption + "']"))).MoveByOffset(2, 1).Perform();
                if (windowOption.ToLower().Equals(selectOption.ToLower()))
                    windowButton = driver.FindElement(By.XPath("//div[@id='" + chartHeaderId + "']//ul[@data-headeroption-id='0']//li/a[text()='" + windowOption + "']"));
            }

            if (selectOption != "")
            {
                Assert.AreNotEqual(null, windowButton, "'" + selectOption + "' not found.");
                windowButton.Click();
                Thread.Sleep(10000);
            }

            Results.WriteStatus(test, "Pass", "Verified, Help Icon on Pages By Channel or Parent Retailer Section");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify When Column Is Drilled Down On From Pages by Channel or Parent Retailer
        ///</summary>
        ///<returns></returns>
        public RetailerActivity VerifyWhenColumnIsDrilledDownOnFromPagesByChannel_ParentRetailer()
        {
            string chartId = "";
            if (driver._isElementPresent("xpath", "//div[@id='PageColumnChannelCustChartHeader']"))
                chartId = "PageColumnChannelCustChart0chart";
            else
                chartId = "PageColumnTradeCustChart0chart";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span"), "x Axis not present in Column Chart.");
            IList<IWebElement> xAxisColl = driver._findElements("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span");
            string columnName = xAxisColl[0].Text;
            columnName = columnName.TrimStart(' ');

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//*[name()='rect' and not(@x=0) and not(@y=0)]"), "Bar Graphs not present.");
            IList<IWebElement> barCollection = driver._findElements("xpath", "//div[@id='" + chartId + "']//*[name()='rect' and not(@x=0) and not(@y=0)]");
            barCollection[0].Click();
            Thread.Sleep(2000);

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='PageColumnChannelCustChartHeader']//li[text()]"), "New Header not present.");
            Assert.AreEqual(columnName.ToLower() + " - retailer", driver._getText("xpath", "//div[@id='PageColumnChannelCustChartHeader']//li[text()]").ToLower(), "Column Name is not updated in the header.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span"), "x Axis not present in Column Chart.");
            xAxisColl = driver._findElements("xpath", "//div[@id='" + chartId + "']//div[@class='highcharts-axis-labels highcharts-xaxis-labels']/span");
            Assert.AreNotEqual(columnName, xAxisColl[0].Text, "Column Chart not updated.");

            Results.WriteStatus(test, "Pass", "Verified, When Column Is Drilled Down On From Pages by Channel or Parent Retailer");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Ad Details Popup
        ///</summary>
        ///<param name="clickCancel">Whether to click Cancel button</param>
        ///<returns></returns>
        public RetailerActivity VerifyAdDetailsPopup(string headerCol, bool clickCancel = false)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@id, 'popupDiv')]"), "Ad Details popup not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@id, 'dvTitle')]"), "Ad Details popup header not present.");
            Assert.IsTrue(driver._getText("xpath", "//div[contains(@id, 'dvTitle')]").Contains(headerCol), "Ad Details popup Header Text not correct.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@id, 'popupDiv')]//img"), "Image not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'popupDiv')]//input"), "Cancel button not present.");

            if (clickCancel)
            {
                driver._click("xpath", "//div[contains(@id, 'popupDiv')]//input");
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[contains(@id, 'popupDiv')]"), "Ad Details pop not closed.");
                Results.WriteStatus(test, "Pass", "Verified, Ad Details Popup is closed");
            }

            Results.WriteStatus(test, "Pass", "Verified, Ad Details Popup");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Open Ad Details Popup
        ///</summary>
        ///<returns></returns>
        public string OpenAdDetailsPopup()
        {
            string headColumn = "", result = "";
            IList<IWebElement> headColumnColl = null;
            IList<IWebElement> clickColumnColl = null;
            if (driver._isElementPresent("xpath", "//table[@id='dvAccountMarketPageGrid_tblMain']"))
            {
                headColumnColl = driver._findElements("xpath", "//table[@id='dvAccountMarketPageGrid_tblMain']//tr[not(@class)]/td[not(@onclick)]");
                clickColumnColl = driver._findElements("xpath", "//table[@id='dvAccountMarketPageGrid_tblMain']//tr[not(@class)]/td[@onclick]");
                result = "Opened, Ad Details Popup from Ads by Week";
            }
            else if (driver._isElementPresent("xpath", "//table[@id='dvAccountMarketFlyerGrid_tblMain']"))
            {
                headColumnColl = driver._findElements("xpath", "//table[@id='dvAccountMarketFlyerGrid_tblMain']//tr[not(@class)]/td[not(@onclick)]");
                clickColumnColl = driver._findElements("xpath", "//table[@id='dvAccountMarketFlyerGrid_tblMain']//tr[not(@class)]/td[@onclick]");
                result = "Opened, Ad Details Popup from Pages by Week";
            }
            else if (driver._isElementPresent("xpath", "//table[@id='dvAccountMarketEventGrid_tblMain']"))
            {
                headColumnColl = driver._findElements("xpath", "//table[@id= 'dvAccountMarketEventGrid_tblMain']//tr[not(@class)]/td[not(@onclick)]");
                clickColumnColl = driver._findElements("xpath", "//table[@id= 'dvAccountMarketEventGrid_tblMain']//tr[not(@class)]/td[@onclick]");
                result = "Opened, Ad Details Popup from Events by Week";
            }
            else if (driver._isElementPresent("xpath", "//table[@id='dvAccountMarketEventGrid_tblMain']"))
            {
                headColumnColl = driver._findElements("xpath", "//table[@id= 'dvAccountMarketEventGrid_tblMain']//tr[not(@class)]/td[not(@onclick)]");
                clickColumnColl = driver._findElements("xpath", "//table[@id= 'dvAccountMarketEventGrid_tblMain']//tr[not(@class)]/td[@onclick]");
                result = "Opened, Ad Details Popup from Themes by Week";
            }

            Assert.IsTrue(headColumnColl != null && clickColumnColl != null, "Tables not found.");

            headColumn = headColumnColl[0].Text;
            clickColumnColl[0].Click();

            Results.WriteStatus(test, "Pass", result);
            return headColumn;
        }

        ///<summary>
        ///Verify Ads/Pages By Week Section
        ///</summary>
        ///<returns></returns>
        public RetailerActivity VerifyAds_PagesByWeekSection(bool adByWeek = true)
        {
            string headerId = "", id = "", pageName = "", result = "";

            if (adByWeek)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Ads']"), "Ads Radio Button not present.");
                driver._click("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Ads']");
                Thread.Sleep(1000);
                headerId = "dvAccountMarketFlyerGridHeader";
                pageName = "Ads";
                id = "dvAccountMarketFlyerGrid";
                result = "Verified, Ads By Week Section";
            }
            else
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Pages']"), "Pages Radio Button not present.");
                driver._click("xpath", "//div[@class='tab-pane ng-scope active']//li[text()='Pages']");
                Thread.Sleep(1000);
                pageName = "Pages";
                headerId = "dvAccountMarketPageGridHeader";
                id = "dvAccountMarketPageGrid";
                result = "Verified, Pages By Week Section";
            }

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='" + headerId + "']//span/div/div[text()]"), "'" + pageName + " by Week' Header not present.");
            Assert.AreEqual(pageName.ToLower() + " by week", driver._getText("xpath", "//div[@id='" + headerId + "']//span/div/div[text()]").ToLower(), "'" + pageName + " by Week' Header text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + headerId + "']//span//div[text() and not(@style)]"), "'" + pageName + " by Week' Drops Header not present.");
            string[] dropList = new string[] { "No Drop", "1 Drop", "2 Drops", "3 Drops", "4 or More Drops" };
            IList<IWebElement> dropNameColl = driver._findElements("xpath", "//div[@id='" + headerId + "']//span//div[text() and not(@style)]");

            foreach (string dropName in dropList)
            {
                bool avail = false;
                foreach (IWebElement drop in dropNameColl)
                    if (drop.Text.ToLower().Contains(dropName.ToLower()))
                    {
                        avail = true;
                        break;
                    }
                Assert.IsTrue(avail, "'" + dropName + "' not found.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + headerId + "']//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + headerId + "']//div[@title='Export']"), "Export Icon not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + id + "']//table"), "Table not present.");

            Results.WriteStatus(test, "Pass", result);
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Events/Themes By Week Section
        ///</summary>
        ///<param name="eventsByWeek">Whether to </param>
        ///<returns></returns>
        public RetailerActivity VerifyEvents_ThemesByWeekSection(bool eventsByWeek = true)
        {
            string headerId = "dvAccountMarketEventGridHeader", id = "dvAccountMarketEventGrid", pageName = "", result = "";
            string[] dropList = null;
            if (eventsByWeek)
            {
                pageName = "Events";
                dropList = new string[] { "First appearance of event", "Second appearance of event", "No Event", "No Drop" };
                result = "Verified, Events By Week Section";
            }
            else
            {
                pageName = "Themes";
                dropList = new string[] { "First appearance of theme", "Second appearance of theme", "No Theme", "No Drop" };
                result = "Verified, Themes By Week Section";
            }

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='" + headerId + "']//span/div/div[text()]"), "'" + pageName + " by Week' Header not present.");

            if (eventsByWeek)
                Assert.AreEqual(pageName.ToLower() + " by week", driver._getText("xpath", "//div[@id='" + headerId + "']//span/div/div[text()]").ToLower(), "'" + pageName + " by Week' Header text does not match.");
            else
                Assert.AreEqual("front page theme usage by week", driver._getText("xpath", "//div[@id='" + headerId + "']//span/div/div[text()]").ToLower(), "'" + pageName + " by Week' Header text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + headerId + "']//table//td[text()]"), "'" + pageName + " by Week' Drops Header not present.");
            IList<IWebElement> dropNameColl = driver._findElements("xpath", "//div[@id='" + headerId + "']//table//td[text()]");

            foreach (string dropName in dropList)
            {
                bool avail = false;
                foreach (IWebElement drop in dropNameColl)
                    if (drop.Text.ToLower().Contains(dropName.ToLower()))
                    {
                        avail = true;
                        break;
                    }
                Assert.IsTrue(avail, "'" + dropName + "' not found.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + headerId + "']//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + headerId + "']//div[@title='Export']"), "Export Icon not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + id + "']//table"), "Table not present.");

            Results.WriteStatus(test, "Pass", result);
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Help Icon on Ads/Pages By Week Section
        ///</summary>
        ///<returns></returns>
        public RetailerActivity VerifyHelpIconOnAds_PagesByWeekSection()
        {
            string chartHeaderId = "";
            if (driver._isElementPresent("xpath", "//div[@id='dvAccountMarketFlyerGridHeader']"))
                chartHeaderId = "dvAccountMarketFlyerGridHeader";
            else
                chartHeaderId = "dvAccountMarketPageGridHeader";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + chartHeaderId + "']//div[@class='fa fa-question homeChartSearch']"), "Help Button not present.");
            driver.MouseHoverUsingElement("xpath", "//div[@id='" + chartHeaderId + "']//div[@class='fa fa-question homeChartSearch']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='" + chartHeaderId + "']//div[@onmouseout='HideTooltip();']"), "Help Icon Tooltip not present.");
            string tooltipText = "<b>How many flyers are my retailers dropping each week? Are retailers increasing or decreasing their frequency over time?</b><br/><br/>Provides a total number of unique Circulars by retailer/market by week. Ads are defined as: Circular, FSI, Mailer, Newspaper  Media Types only. If ‘Product’ level selections are made, it will ignore these selections and return ad level results, i.e. total number of ads/pages in a Circular, regardless of categories advertised. Color coding indicates the number of ads i.e. two ad drops will be yellow.";
            Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@id='" + chartHeaderId + "']//div[@onmouseout='HideTooltip();']", "onmouseover").Contains(tooltipText), "Help Icon Text does not match.");

            Results.WriteStatus(test, "Pass", "Verified, Help Icon on Ads/Pages By Week Section");
            return new RetailerActivity(driver, test);
        }

        ///<summary>
        ///Verify Help Icon on Events/Themes By Week Section
        ///</summary>
        ///<returns></returns>
        public RetailerActivity VerifyHelpIconOnEvents_ThemesByWeekSection()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvAccountMarketEventGridHeader']//div[@class='fa fa-question homeChartSearch']"), "Help Button not present.");
            driver.MouseHoverUsingElement("xpath", "//div[@id='dvAccountMarketEventGridHeader']//div[@class='fa fa-question homeChartSearch']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvAccountMarketEventGridHeader']//div[@onmouseout='HideTooltip();']"), "Help Icon Tooltip not present.");
            string tooltipText1 = "<b>When are retailers beginning to advertise for key events? Are they promoting a holiday early in the season or later?</b><br/><br/>Provides a summary of Events at an ad block level for the categories you subscribe to. (For front page ad activity narrow your results to ‘Front’ page position.) Color coding i.e. green, indicates the first appearance of the event. Events are tied to the ads, and include any holiday theming or special events like Super Bowl.";
            string tooltipText2 = "<b>When are retailers beginning to advertise for key events or themes? Are they promoting a holiday early in the season or later?</b><br/><br/>Provides a summary of Front Page Events or Themes. Color coding i.e. green, indicates the first appearance of the event. Events are tied to the ads, and include things like Dollar Sale and Super Sale. Themes include any holiday theming or special events like Super Bowl.";
            Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@id='dvAccountMarketEventGridHeader']//div[@onmouseout='HideTooltip();']", "onmouseover").Contains(tooltipText1)
                || driver._getAttributeValue("xpath", "//div[@id='dvAccountMarketEventGridHeader']//div[@onmouseout='HideTooltip();']", "onmouseover").Contains(tooltipText2),
                "Help Icon Text does not match.");

            Results.WriteStatus(test, "Pass", "Verified, Help Icon on Events/Themes By Week Section");
            return new RetailerActivity(driver, test);
        }




        #endregion
    }
}
