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
    public class PricingHistograms
    {
        #region Private Variables

        private IWebDriver pricingHistograms;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public PricingHistograms(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.pricingHistograms = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.pricingHistograms; }
            set { this.pricingHistograms = value; }
        }

        ///<summary>
        ///Verify Pricing Histograms Page
        ///</summary>
        ///<param name="clientName">Client Name</param>
        ///<param name="fromPricingPromotions">Whether navigating from pricing & promotions screen</param>
        ///<param name="searchName">Name of Saved Search loaded on screen</param>
        ///<returns></returns>
        public PricingHistograms VerifyPricingHistogramsPage(bool fromPricingPromotions = true, string clientName = "Procter & Gamble", string searchName = "")
        {
            if (fromPricingPromotions)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//h1[text()='Pricing & Promotions']"), "Category & Brand Share Screen not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='title cursorpointer']/a[text()='Pricing Histograms']"), "Pricing Histograms link not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-media']/i[contains(@class, 'promo-report-generic')]"), "Pricing Histograms icon not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-content']/div[contains(@class, 'submenudescription')]"), "Pricing Histograms description not present.");

                driver._click("xpath", "//div[@class='title cursorpointer']/a[text()='Pricing Histograms']");
                Thread.Sleep(5000);
            }

            driver._waitForElementToBeHidden("xpath", "//div[@class='ProcessLoader'], 45");

            Assert.IsTrue(driver._waitForElement("xpath", "//img[@src='/Images/isc-NumeratorLogo.png']", 15), "Numerator Logo not found on Home Page.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='report-content']/h1"), "Madlib Search header text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[contains(@class, 'filler')]"), "Madlib Prompt Summary text not present.");

            if (searchName != "")
                Assert.IsTrue(driver._getText("xpath", "//div[@class='report-content']/h1").ToLower().Contains(searchName.ToLower()), "Search Name '" + searchName + "' does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='1']"), "Madlib Search Parameter 'Any Product' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='2']"), "Madlib Search Parameter 'Any Account' not present.");
            if (!clientName.ToLower().Contains("australia"))
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='3']"), "Madlib Search Parameter 'Any Market' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='4']"), "Madlib Search Parameter 'Any Date' not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Save Options']/i"), "'Save' option not present.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@role='tabpanel']//li[@heading]"), "Tabs not present.");
            string[] tabNameList = new string[] { "Channel", "Retailer" };
            IList<IWebElement> tabCollection = driver._findElements("xpath", "//div[@role='tabpanel']//li[@heading]");

            foreach (string tabName in tabNameList)
            {
                bool avail = false;
                foreach (IWebElement tab in tabCollection)
                    if (tab.GetAttribute("heading").ToLower().Contains(tabName.ToLower()))
                    {
                        avail = true;
                        break;
                    }
                Assert.IsTrue(avail, "'" + tabName + "' tab not found.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Ad Sharing And Exclusivity Page");
            return new PricingHistograms(driver, test);
        }

        ///<summary>
        ///Verify Channel/Retailer Tab
        ///</summary>
        ///<param name="tabName">Tab to be selected</param>
        ///<returns></returns>
        public PricingHistograms VerifyChannel_RetailerTab(string tabName = "Channel")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@role='tabpanel']//li[@heading='" + tabName + "']"), "" + tabName + " Tab not present.");
            if(!driver._getAttributeValue("xpath", "//div[@role='tabpanel']//li[@heading='" + tabName + "']", "class").Contains("active"))
            {
                driver._click("xpath", "//div[@role='tabpanel']//li[@heading='" + tabName + "']");
                Thread.Sleep(2000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@role='tabpanel']//li[@heading='" + tabName + "']", "class").Contains("active"), "" + tabName + " tab not selected.");
            }

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//span[@object-type='pagetitle']"), "'Page Header not present.'");
            Assert.IsTrue(driver._getText("xpath", "//div[contains(@class, 'active')]//span[@object-type='pagetitle']").Contains("Advertised Price Point by Category"), "'Page Header' text does not match.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//div[contains(@class, 'suggestionText')]//span"), "'Info Text' not present.");
            Assert.AreEqual("*(Click on any data point to view product images)", driver._getText("xpath", "//div[contains(@class, 'active')]//div[contains(@class, 'suggestionText')]//span"), "'Info Text' does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//*[name()='svg']"), "Charts not present.");
            if(tabName.Equals("Channel"))
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@class, 'legendText')]"), "Legend not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//*[contains(@id,'hartContainer0Header')]"), "Chart Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@class, 'question homeChartSearch')]"), "Help Icon not present in charts");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@title='More Options']"), "Window Icon not present in charts.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@title='Export']"), "Export Icon not present.");

            Results.WriteStatus(test, "Pass", "Verified, " + tabName + " Tab");
            return new PricingHistograms(driver, test);
        }

        ///<summary>
        ///Verify Pricing Histograms Help Icon
        ///</summary>
        ///<returns></returns>
        public PricingHistograms VerifyPricingHistogramsPageHelpIcon()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id,'hartContainer0Header')]//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            driver._scrollintoViewElement("xpath", "//div[contains(@class, 'active')]//div[contains(@id,'hartContainer0Header')]//div[@class='fa fa-question homeChartSearch']");
            driver.MouseHoverUsingElement("xpath", "//div[contains(@class, 'active')]//div[contains(@id,'hartContainer0Header')]//div[@class='fa fa-question homeChartSearch']");

            string capturedHelpIconText = driver._getAttributeValue("xpath", "//div[contains(@class, 'active')]//div[contains(@id,'hartContainer0Header')]//div[@onmouseout='HideTooltip();']", "onmouseover");
            string expectedHelpIconText1 = "Pricing Histograms display the advertised price point by category providing visibility into the frequency of a unique promoted price point based on the 5 selected retail competitors.";
            string expectedHelpIconText2 = "Red Dotted Line = Median Unit Price for Target(middle value in the list of price points)";
            string expectedHelpIconText3 = "Red Dashed Line = Mean Unit Price for Target(average - add up all the price points then divide by the number of price points)";
            string expectedHelpIconText4 = "Absence of red lines can signify:";
            string expectedHelpIconText5 = "Items are promoted with no unit price points";
            string expectedHelpIconText6 = "There are no promotions for 'Me' Retailer";
            string expectedHelpIconText7 = "'Me' Retailer is not included in my search parameters";

            Console.WriteLine(capturedHelpIconText);
            Assert.IsTrue(capturedHelpIconText.ToLower().Contains(expectedHelpIconText1.ToLower()) 
                && capturedHelpIconText.ToLower().Contains(expectedHelpIconText2.ToLower())
                && capturedHelpIconText.ToLower().Contains(expectedHelpIconText3.ToLower())
                && capturedHelpIconText.ToLower().Contains(expectedHelpIconText4.ToLower())
                && capturedHelpIconText.ToLower().Contains(expectedHelpIconText5.ToLower())
                && capturedHelpIconText.ToLower().Contains(expectedHelpIconText6.ToLower())
                && capturedHelpIconText.ToLower().Contains(expectedHelpIconText7.ToLower()), "Help Icon tooltip text did not match.");

            Results.WriteStatus(test, "Pass", "Verified, Help Icon for 'Manufacturer Comparison' Page.");
            return new PricingHistograms(driver, test);
        }

        ///<summary>
        ///Verify Windows Icon on Channel Tab Charts
        ///</summary>
        ///<param name="chartName">Name of chart to verify Windows icon on</param>
        ///<param name="optionName">Option to be clicked</param>
        ///<returns></returns>
        public PricingHistograms VerifyWindowsIconOnChannelTabCharts(string chartName = "", string optionName = "")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@class= 'splineChartContainer']"), "Charts not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@title='More Options']"), "Window Icon not present in charts.");
            IList<IWebElement> chartsCollection = driver._findElements("xpath", "//div[contains(@class, 'active')]//div[@class= 'splineChartContainer']");

            IWebElement windowIcon = null;
            int chartIndex = 0;
            if(chartName != "")
            {
                foreach(IWebElement chart in chartsCollection)
                {
                    IList<IWebElement> chartTitleColl = chart._findElementsWithinElement("xpath", ".//ul[@class='prompt-breadcrumbs-list']//span");
                    if (chartTitleColl[0].Text.ToLower().Contains(chartName.ToLower()))
                        break;
                    ++chartIndex;
                }
            }

            IList<IWebElement> windowIconColl = chartsCollection[chartIndex]._findElementsWithinElement("xpath", ".//div[@title='More Options']");
            windowIcon = windowIconColl[0];
            Assert.AreNotEqual(null, windowIcon, "Window icon not found on chart.");

            windowIcon.Click();
            IList<IWebElement> ddlOptions = chartsCollection[chartIndex]._findElementsWithinElement("xpath", ".//ul[@data-headeroption-id=0]/li");
            Assert.AreNotEqual(0, ddlOptions.Count, "DDL is not present.");

            Actions actions = new Actions(driver);
            bool avail = false;
            foreach (IWebElement option in ddlOptions)
            {
                actions.MoveToElement(option).Perform();
                if (option.Text.ToLower().Equals(optionName.ToLower()))
                {
                    avail = true;
                    option.Click();
                    break;
                }
            }
            Assert.IsTrue(avail, "'" + optionName + "' not found.");
            Thread.Sleep(2000);

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@class= 'splineChartContainer']"), "Charts not present.");
            chartsCollection = driver._findElements("xpath", "//div[contains(@class, 'active')]//div[@class= 'splineChartContainer']");

            IList<IWebElement> yaxisCollection = chartsCollection[chartIndex]._findElementsWithinElement("xpath", ".//div[contains(@class, 'highcharts-yaxis-labels')]/span");
            Assert.AreNotEqual(0, yaxisCollection.Count, "Y-Axis values not present");

            if (optionName.Contains("%"))
            {
                foreach (IWebElement yaxis in yaxisCollection)
                    Assert.IsTrue(yaxis.Text.Contains("%"), "'Show by %' not applied successfully.");
                Results.WriteStatus(test, "Pass", "'Show by %' is applied successfully.");
            }
            else
            {
                foreach (IWebElement yaxis in yaxisCollection)
                    Assert.IsFalse(yaxis.Text.Contains("%"), "'Show by #' not applied successfully.");
                Results.WriteStatus(test, "Pass", "'Show by #' is applied successfully.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Windows Icon on Channel Tab Charts");
            return new PricingHistograms(driver, test);
        }

        ///<summary>
        ///Verify Export Icon on Channel Tab Charts
        ///</summary>
        ///<param name="chartName">Name of chart to verify Windows icon on</param>
        ///<param name="optionName">Option to be clicked</param>
        ///<returns></returns>
        public string VerifyExportIconOnChannelTabCharts(string chartName = "", string optionName = "")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@class= 'splineChartContainer']"), "Charts not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@title='Export']"), "Export Icon not present in charts.");
            IList<IWebElement> chartsCollection = driver._findElements("xpath", "//div[contains(@class, 'active')]//div[@class= 'splineChartContainer']");

            IWebElement exportIcon = null;
            int chartIndex = 0;
            IList<IWebElement> chartTitleColl = null;
            if (chartName != "")
            {
                foreach (IWebElement chart in chartsCollection)
                {
                    chartTitleColl = chart._findElementsWithinElement("xpath", ".//ul[@class='prompt-breadcrumbs-list']//span");
                    if (chartTitleColl[0].Text.ToLower().Contains(chartName.ToLower()))
                        break;
                    ++chartIndex;
                }
            }

            chartTitleColl = chartsCollection[chartIndex]._findElementsWithinElement("xpath", ".//ul[@class='prompt-breadcrumbs-list']//span");
            chartName = chartTitleColl[0].Text;
            IList<IWebElement> exportIconColl = chartsCollection[chartIndex]._findElementsWithinElement("xpath", ".//div[@title='Export']");
            exportIcon = exportIconColl[0];
            Assert.AreNotEqual(null, exportIcon, "Window icon not found on chart.");

            exportIcon.Click();
            IList<IWebElement> ddlOptions = chartsCollection[chartIndex]._findElementsWithinElement("xpath", ".//ul[@data-headeroption-id=1]/li");
            Assert.AreNotEqual(0, ddlOptions.Count, "DDL is not present.");

            Actions actions = new Actions(driver);
            bool avail = false;
            foreach (IWebElement option in ddlOptions)
            {
                actions.MoveToElement(option).Perform();
                if (option.Text.ToLower().Equals(optionName.ToLower()))
                {
                    avail = true;
                    option.Click();
                    break;
                }
            }
            Assert.IsTrue(avail, "'" + optionName + "' not found.");
            Thread.Sleep(15000);

            if (chartName.Contains("&"))
                chartName = chartName.Replace("&", "and");

            Results.WriteStatus(test, "Pass", "Verified, Export Icon on Channel Tab Charts");
            return chartName;
        }

        ///<summary>
        ///Verify Data Point Popup
        ///</summary>
        ///<param name="chartName">Name of chart</param>
        ///<param name="clickCancel">Whether popup should be visible of not</param>
        ///<returns></returns>
        public string VerifyDataPointPopup(bool clickCancel = false, string chartName = "")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@class= 'splineChartContainer']"), "Charts not present.");
            IList<IWebElement> chartsCollection = driver._findElements("xpath", "//div[contains(@class, 'active')]//div[@class= 'splineChartContainer']");

            int chartIndex = 0;
            IList<IWebElement> chartTitleColl = null;
            if (chartName != "")
            {
                foreach (IWebElement chart in chartsCollection)
                {
                    chartTitleColl = chart._findElementsWithinElement("xpath", ".//ul[@class='prompt-breadcrumbs-list']//span");
                    if (chartTitleColl[0].Text.ToLower().Contains(chartName.ToLower()))
                        break;
                    ++chartIndex;
                }
            }

            chartTitleColl = chartsCollection[chartIndex]._findElementsWithinElement("xpath", ".//ul[@class='prompt-breadcrumbs-list']//span");
            chartName = chartTitleColl[0].Text;
            IList<IWebElement> legendColl = chartsCollection[chartIndex]._findElementsWithinElement("xpath", ".//*[name()='g' and contains(@class, 'legend-item')]//*[name()='tspan']");
            string legend = legendColl[0].Text;
            IList<IWebElement> dataPointColl = chartsCollection[chartIndex]._findElementsWithinElement("xpath", ".//*[name()='path' and not(@fill='none')]");
            Assert.AreNotEqual(0, dataPointColl.Count, "Data Points not present on chart '" + chartName + "'");

            int x = 0;
            Random rand = new Random();
            if (dataPointColl.Count > 5)
                x = rand.Next(1, 5);
            else
                x = rand.Next(0, dataPointColl.Count);

            Actions action = new Actions(driver);
            action.MoveToElement(dataPointColl[x-1]).MoveToElement(dataPointColl[x]).Perform();

            Assert.IsTrue(driver._waitForElement("xpath", "//*[contains(@class, 'tooltip')]//td[contains(text(), '$')]"), "Data Point Unit Price not present in popup.");
            string unitPrice = driver._getText("xpath", "//*[contains(@class, 'tooltip')]//td[contains(text(), '$')]");
            action.MoveToElement(dataPointColl[x-1]).MoveToElement(dataPointColl[x]).Click().Perform();
            Thread.Sleep(2000);

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'modal-content')]"), "Data Point popup not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'modal-header')]//h4"), "Data Point popup header not present.");

            Assert.IsTrue(driver._getText("xpath", "//div[contains(@class, 'modal-header')]//h4").Contains("Drill Detail(s):")
                && driver._getText("xpath", "//div[contains(@class, 'modal-header')]//h4").Contains("Category")
                && driver._getText("xpath", "//div[contains(@class, 'modal-header')]//h4").Contains(chartName)
                && driver._getText("xpath", "//div[contains(@class, 'modal-header')]//h4").Contains("Retailer:")
                && driver._getText("xpath", "//div[contains(@class, 'modal-header')]//h4").Contains(legend)
                && driver._getText("xpath", "//div[contains(@class, 'modal-header')]//h4").Contains("Unit Price:")
                && driver._getText("xpath", "//div[contains(@class, 'modal-header')]//h4").Contains(unitPrice), "Data Point popup header text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'detailHeader')]//span[contains(@class, 'viewModeText')]"), "'Detail Data' Header not present.");
            Assert.AreEqual("Detail Data", driver._getText("xpath", "//div[contains(@class, 'detailHeader')]//span[contains(@class, 'viewModeText')]"), "'Detail Data' Header text does not match.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'chkLbl') and text()='Detail Data']"), "'Detail Data' Radio button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'chkLbl') and text()='Promoted Product Images']"), "'Promoted Product Images' Radio button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@class, 'chkLbl') and text()='Page Images']"), "'Page Images' Radio button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'detailHeader')]//div[@class='CustChartfilter-header']"), "'Export' icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'pageSizeGrp')]/button"), "'Records per Page' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//ul[contains(@class, 'pagination')]//a"), "'Page Navigation' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'modal-footer')]//button[text()='Cancel']"), "Cancel button not present on popup.");

            Results.WriteStatus(test, "Pass", "Verified, Data Point Popup.");

            if (clickCancel)
            {
                Thread.Sleep(5000);
                driver._click("xpath", "//div[contains(@class, 'modal-footer')]//button[text()='Cancel']");
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'modal-content')]"), "Data Popup is not closed.");
                Results.WriteStatus(test, "Pass", "Verified, Data Point Popup is closed.");
            }

            return chartName = "";
        }

        #endregion
    }
}
