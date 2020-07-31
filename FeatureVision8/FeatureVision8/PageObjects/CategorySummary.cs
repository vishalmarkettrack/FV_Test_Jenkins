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
    public class CategorySummary
    {
        #region Private Variables

        private IWebDriver categorySummary;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public CategorySummary(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.categorySummary = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.categorySummary; }
            set { this.categorySummary = value; }
        }

        ///<summary>
        ///Verify Category Summary Page
        ///</summary>
        ///<param name="clientName">Name of the client</param>
        ///<param name="searchName">Searchname to be verified</param>
        ///<returns></returns>
        public CategorySummary VerifyCategorySummaryPage(string clientName = "Target", string searchName = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//h1[text()='Category & Brand Share']"), "Category & Brand Share Screen not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='title cursorpointer']/a[text()='Category Summary']"), "Category Summary link not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-media']/i[contains(@class, 'category-summary')]"), "Category Summary icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-content']/div[contains(@class, 'submenudescription')]"), "Category Summary description not present.");

            driver._click("xpath", "//div[@class='title cursorpointer']/a[text()='Category Summary']");
            Thread.Sleep(1000);

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

            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Save Options']/i"), "'Search' option not present.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@role='tabpanel']//li[@heading]"), "Tabs not present.");
            string[] tabNameList = new string[] { "Promoted Products", "Top 10 Manufacturer", "Ad Type by Retailers" };
            if (clientName.ToLower().Contains("target"))
            {
                tabNameList[0] = "Top 10 Retailer";
                tabNameList[1] = "Promoted Products by Manufacturer";
            }
            else if (clientName.ToLower().Contains("jcpenney"))
            {
                tabNameList[0] = "Top 10 Brand";
                tabNameList[1] = "Promoted Products by Brand";
            }
            else if (clientName.ToLower().Contains("metcash"))
            {
                tabNameList[0] = "Top 10 Retailer";
                tabNameList[1] = "Promoted Products by Manufacturer/ Distributor";
                tabNameList[2] = "Medium by Retailers";
            }
            else if (clientName.ToLower().Contains("activision"))
            {
                tabNameList[0] = "Promoted Products";
                tabNameList[1] = "Top 10 Brand";
            }
            else if (clientName.ToLower().Contains("mattel"))
            {
                tabNameList[0] = "Promoted Products";
                tabNameList[1] = "Top 10 Manufacturer/ Distributor";
                tabNameList[2] = "Medium by Retailers";
            }
            else if (clientName.ToLower().Contains("california table grape"))
            {
                Array.Resize(ref tabNameList, 4);
                tabNameList[0] = "Promoted Product by Retailer";
                tabNameList[1] = "Promoted Product by Origin";
                tabNameList[2] = "Promoted Product by Variety";
                tabNameList[3] = "Ad Type by Retailers";
            }
            else if (clientName.ToLower().Contains("meyer corporation"))
            {
                Array.Resize(ref tabNameList, 4);
                tabNameList[0] = "Promoted Product by Retailer Group";
                tabNameList[1] = "Promoted Product by Manufacturer";
                tabNameList[2] = "Promoted Product by Category";
                tabNameList[3] = "Ad Type by Retailers";
            }

            else if (clientName.ToLower().Contains("logitech"))
            {
                Array.Resize(ref tabNameList, 4);
                tabNameList[0] = "Promoted Product by Channel";
                tabNameList[1] = "Promoted Product by Manufacturer/ Distributor";
                tabNameList[2] = "Promoted Product by Subcategory";
                tabNameList[3] = "Medium by Retailers";
            }

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

            Results.WriteStatus(test, "Pass", "Verified, Category Summary Page");
            return new CategorySummary(driver, test);
        }

        ///<summary>
        ///Select Tab On Category Summary Page
        ///</summary>
        ///<param name="tabName">Tab to be selected</param>
        ///<returns></returns>
        public CategorySummary SelectTabOnCategorySummaryPage(string tabName)
        {
            IList<IWebElement> tabCollection = driver._findElements("xpath", "//div[@role='tabpanel']//li[@heading]");

            bool avail = false;
            foreach (IWebElement tab in tabCollection)
                if (tab.GetAttribute("heading").ToLower().Contains(tabName.ToLower()))
                {
                    avail = true;
                    if (!tab.GetAttribute("class").Contains("active"))
                        tab.Click();
                    break;
                }
            Assert.IsTrue(avail, "'" + tabName + "' tab not found.");

            Thread.Sleep(2000);

            Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@role='tabpanel']//li[@heading='" + tabName + "']", "class").Contains("active"), "'" + tabName + "' Tab not selected.");

            Results.WriteStatus(test, "Pass", "Selected, '" + tabName + "' Tab On Category Summary Page");
            return new CategorySummary(driver, test);
        }

        ///<summary>
        ///Verify Category Summary Help Icon
        ///</summary>
        ///<param name="tabName">Tab Name for Help Icon</param>
        ///<returns></returns>
        public CategorySummary VerifyCategorySummaryPageHelpIcon(string tabName = "Top 10 Retailer")
        {
            string capturedHelpIconText = "";
            if (tabName.ToLower().Equals("promoted products"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'active')]//div[@ng-show and @class='']//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
                driver._scrollintoViewElement("xpath", "//div[contains(@class,'active')]//div[@ng-show and @class='']//div[@class='fa fa-question homeChartSearch']");
                driver.MouseHoverUsingElement("xpath", "//div[contains(@class,'active')]//div[@ng-show and @class='']//div[@class='fa fa-question homeChartSearch']");

                capturedHelpIconText = driver._getAttributeValue("xpath", "//div[contains(@class,'active')]//div[@ng-show and @class='']//div[@onmouseout='HideTooltip();']", "onmouseover");
            }
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'active')]//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
                driver._scrollintoViewElement("xpath", "//div[contains(@class,'active')]//div[@class='fa fa-question homeChartSearch']");
                driver.MouseHoverUsingElement("xpath", "//div[contains(@class,'active')]//div[@class='fa fa-question homeChartSearch']");

                capturedHelpIconText = driver._getAttributeValue("xpath", "//div[contains(@class,'active')]//div[@onmouseout='HideTooltip();']", "onmouseover");
            }

            string expectedHelpIconText1 = "";
            string expectedHelpIconText2 = "";

            if (tabName.ToLower().Equals("top 10 retailer") || tabName.ToLower().Equals("top 10 brand"))
            {
                expectedHelpIconText1 = "Do we have the highest (or lowest) share of the category?";
                expectedHelpIconText2 = "Provides a total of promoted products in a given period, compared to last period (if selected). The comparison is between me (retailer) vs the top competitors. The chart will also show the percentage of gain/loss (if last period is selected). Tip: Narrow your search to a single key category.";
            }
            else if (tabName.ToLower().Equals("promoted products by manufacturer") || tabName.ToLower().Equals("promoted products by brand"))
            {
                expectedHelpIconText1 = "Are Manufacturers promoting more heavily wrt/competition? Did we see an increase (or decrease) in Manufacturer promotions?";
                expectedHelpIconText2 = "Provides a total of promoted products by manufacturer in a given period, compared to last period (if selected). The comparison is between me (retailer) vs other retailers of the same category. Drill down (click on any bar in chart) to see additional views. Tip: Narrow your search to a single key category.";
            }
            else if (tabName.ToLower().Equals("promoted products") || tabName.ToLower().Equals("promoted product by retailer") || tabName.ToLower().Equals("promoted product by retailer group") || tabName.ToLower().Equals("promoted products by channel"))
            {
                expectedHelpIconText1 = "Are we promoting heavily with relation to competition? Did we see an increase (or decrease) in Channel or at Retailer?";
                expectedHelpIconText2 = "Provides a total of promoted products by channel (Drug, Mass, Food) in a given period, compared to last period (if selected). The comparison is between me (manufacturer) vs other manufacturers of the same category, with an option to change the value from channel to parent retailer. Drill down (click on any bar in chart) to see additional views.";
            }
            else if (tabName.ToLower().Equals("top 10 manufacturer") || tabName.ToLower().Equals("promoted product by origin") || tabName.ToLower().Equals("promoted product by variety") || tabName.ToLower().Equals("promoted product by category") || tabName.ToLower().Equals("promoted product by subcategory"))
            {
                expectedHelpIconText1 = "What is my share of the category compared to competition?";
                expectedHelpIconText2 = "Provides a total of promoted products in a given period, compared to last period (if selected). The comparison is between me (manufacturer) vs the top 10 competitors. The chart will also show the percentage of gain/loss (if last period is selected). For manufacturer view, click on any bar in the chart to see brands for that manufacturer with the numbers and percentage of gain/loss. Tip: Narrow your search to a single key category.";
            }
            else if (tabName.ToLower().Equals("ad type by retailers") || tabName.ToLower().Equals("medium by retailers"))
            {
                expectedHelpIconText1 = "How prominent are promotions in your category? Which retailers feature your brands more prominently?";
                expectedHelpIconText2 = "Provides a total of promoted products by ad types (A, B, C) in a given period, compared to last period (if selected). The comparison is between retailers. “A” large = Illustrated ads that are larger than other ads on the page. “B” medium = Illustrated ads that are approximately equal in size to other ads on the page. Most ads are assigned an ad type of B. “C” liner = Typically non-illustrated liner listings (no image).";
            }

            Console.WriteLine(capturedHelpIconText);
            Assert.IsTrue(capturedHelpIconText.ToLower().Contains(expectedHelpIconText1.ToLower()) && capturedHelpIconText.ToLower().Contains(expectedHelpIconText2.ToLower()), "Help Icon tooltip text did not match.");

            Results.WriteStatus(test, "Pass", "Verified, Help Icon for '" + tabName + "' Tab.");
            return new CategorySummary(driver, test);
        }

        ///<summary>
        ///Verify Top 10 Tab
        ///</summary>
        ///<param name="tabName">Name of Tab to be Verified</param>
        ///<returns></returns>
        public CategorySummary VerifyTop10Tab(string tabName = "Retailer")
        {
            string id = "Retailer";
            if (tabName.ToLower().Equals("retailer"))
                tabName = "Retailer";
            else if (tabName.ToLower().Equals("brand"))
                tabName = "Brand";
            else if (tabName.ToLower().Equals("manufacturer"))
            {
                tabName = "Manufacturer";
                id = "Manufacture";
            }

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='dvTopTen" + id + "HeaderTitle']//li"), "'Top 10 " + tabName + "' Tab Header Text not present.");
            Assert.AreEqual("Top 10 " + tabName + "", driver._getText("xpath", "//div[@id='dvTopTen" + id + "HeaderTitle']//li/span"), "'Top 10 " + tabName + "' Header Text does not match.");

            if (!tabName.ToLower().Equals("brand"))
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvTopTen" + id + "Header']//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvTopTen" + id + "Header']//div[@title='More Options']"), "'More Options' Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvTopTen" + id + "Header']//div[@title='Export']"), "'Export' Icon not present.");

            if(tabName.ToLower().Equals("manufacturer"))
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvTopTenManufacture0chart']//*[name()='svg']"), "'Bar Chart by " + tabName + "' not present.");
            else
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvTopTenRetailer0chartP']//*[name()='svg']"), "'Bar Chart by " + tabName + "' not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvTopTen" + id + "0chartTabular_tblMain']"), "'Top 10 " + tabName + " and Values' Table not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvTopTen" + id + "0chartTabular_tblMain']//td[@class and text()='Top 10 " + tabName + "']"), "'Top 10 Retailer' Column Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvTopTen" + id + "0chartTabular_tblMain']//td[@class and text()='Values']"), "'Values' Column Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvTopTen" + id + "0chartTabular_tblMain']//td[not(@class)][1]"), "'Top 10 " + tabName + "' Column not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvTopTen" + id + "0chartTabular_tblMain']//td[not(@class)][2]"), "'Values' Column not present.");

            Results.WriteStatus(test, "Pass", "Verified, Top 10 " + tabName + " Tab");
            return new CategorySummary(driver, test);
        }

        ///<summary>
        ///Verify Show by Number Or Percentage On Category Summary Screen
        ///</summary>
        ///<param name="perc">Show by Percentage Or #</param>
        ///<returns></returns>
        public CategorySummary VerifyShowByNumberOrPercentageOnCategorySummaryScreen(bool perc = false)
        {
            if(driver._getAttributeValue("xpath", "//li[@heading and contains(@class, 'active')]", "heading").ToLower().Contains("ad type by retailer") || driver._getAttributeValue("xpath", "//li[@heading and contains(@class, 'active')]", "heading").ToLower().Contains("medium by retailer"))
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//div[contains(@class, 'highcharts-yaxis-labels')]//span"), "Y-Axis Labels not present");
                IList<IWebElement> yAxisLabelColl = driver._findElements("xpath", "//div[contains(@class, 'active')]//div[contains(@class, 'highcharts-yaxis-labels')]//span");

                if (perc)
                {
                    Assert.IsTrue(yAxisLabelColl[0].Text.Contains("%"), "'Show By Percentage' not applied.");
                    Results.WriteStatus(test, "Pass", "Verified, 'Show By Percentage' successfully applied.");
                }
                else
                {
                    Assert.IsFalse(yAxisLabelColl[0].Text.Contains("%"), "'Show By #' not applied.");
                    Results.WriteStatus(test, "Pass", "Verified, 'Show By #' successfully applied.");
                }
            }
            else
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//*[contains(@class, 'labels')]//*[name()='tspan']"), "'Chart Value Labels' not present.");
                IList<IWebElement> chartLabelValuesColl = driver._findElements("xpath", "//div[contains(@class, 'active')]//*[contains(@class, 'labels')]//*[name()='tspan']");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//table[@class='MyCustGrid']//td[not(@class)][2]"), "'Data Table Values' not present.");
                IList<IWebElement> custGridValuesColl = driver._findElements("xpath", "//div[contains(@class, 'active')]//table[@class='MyCustGrid']//td[not(@class)][2]");

                if (driver._isElementPresent("xpath", "//div[contains(@class,'active')]//div[@ng-show and @class='']//div[contains(@id, 'Header')]"))
                {
                    Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//*[contains(@class, 'labels')]//*[name()='tspan']"), "'Chart Value Labels' not present.");
                    chartLabelValuesColl = driver._findElements("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//*[contains(@class, 'labels')]//*[name()='tspan']");

                    Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[@class='MyCustGrid']//td[not(@class)][2]"), "'Data Table Values' not present.");
                    custGridValuesColl = driver._findElements("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[@class='MyCustGrid']//td[not(@class)][2]");
                }

                if (perc)
                {
                    Assert.IsTrue(chartLabelValuesColl[0].Text.Contains("%"), "'Show By Percentage' not applied in chart(s).");
                    Assert.IsTrue(custGridValuesColl[0].Text.Contains("%"), "'Show By Percentage' not applied in table.");
                    Results.WriteStatus(test, "Pass", "Verified, 'Show By Percentage' successfully applied.");
                }
                else
                {
                    Assert.IsFalse(chartLabelValuesColl[0].Text.Contains("%"), "'Show By #' not applied in chart(s).");
                    Assert.False(custGridValuesColl[0].Text.Contains("%"), "'Show By #' not applied in table.");
                    Results.WriteStatus(test, "Pass", "Verified, 'Show By #' successfully applied.");
                }
            }

            return new CategorySummary(driver, test);
        }

        ///<summary>
        ///Verify Promoted Products Tab With Radio Button
        ///</summary>
        ///<param name="tabName">Name of Tab to be Verified</param>
        ///<param name="radioButton">Radio button to be selected</param>
        ///<returns></returns>
        public CategorySummary VerifyPromotedProductsTabWithRadioButton(string tabName = "Promoted Products", string radioButton = "")
        {
            if (!radioButton.Equals(""))
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//li[text()='Channel']"), "'Channel' Radio button not present");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//li[text()='Parent Retailer']"), "'Parent Retailer' Radio button not present");

                if (radioButton.ToLower().Equals("channel"))
                    driver._click("xpath", "//div[contains(@class, 'active')]//li[text()='Channel']");
                else
                    driver._click("xpath", "//div[contains(@class, 'active')]//li[text()='Parent Retailer']");

                Thread.Sleep(2000);
            }

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//div[contains(@id, 'Header')]//li/span"), "'" + tabName + "' Tab Header Text not present.");
            Assert.AreEqual(tabName, driver._getText("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//div[contains(@id, 'Header')]//li/span"), "'" + tabName + "' Header Text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//div[contains(@id, 'Header')]//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//div[contains(@id, 'Header')]//div[@title='More Options']"), "'More Options' Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//div[contains(@id, 'Header')]//div[@title='Export']"), "'Export' Icon not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//div[contains(@id, '0chart')]//*[name()='svg']"), "'Bar Chart by " + tabName + "' not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '0chartTabular')]"), "'" + tabName + " and Values' Table not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '0chartTabular')]//td[@class][1]"), "'" + tabName + "' Column Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '0chartTabular')]//td[@class][2]"), "'Values' Column Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '0chartTabular')]//td[not(@class)][1]"), "'" + tabName + "' Column not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '0chartTabular')]//td[not(@class)][2]"), "'Values' Column not present.");

            if (!tabName.ToLower().Contains("subcategory") || !tabName.ToLower().Contains("category") || !tabName.ToLower().Contains("retailer group") 
                || !tabName.ToLower().Contains("variety") || !tabName.ToLower().Contains("origin") || (tabName.ToLower().Contains("channel") && radioButton == ""))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//div[contains(@id, '1chart')]//*[name()='svg']"), "'Bar Chart by " + tabName + "' not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '1chartTabular')]"), "'Filtered Competition and Values' Table not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '1chartTabular')]//td[@class][1]"), "'Filtered Competition' Column Header not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '1chartTabular')]//td[@class][2]"), "'Values' Column Header not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '1chartTabular')]//td[not(@class)][1]"), "'Filtered Competition' Column not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[@ng-show and @class='']//table[contains(@id, '1chartTabular')]//td[not(@class)][2]"), "'Values' Column not present.");
            }

            Results.WriteStatus(test, "Pass", "Verified, " + tabName + " Tab");
            return new CategorySummary(driver, test);
        }

        ///<summary>
        ///Verify Promoted Products Tab
        ///</summary>
        ///<param name="tabName">Name of Tab to be Verified</param>
        ///<returns></returns>
        public CategorySummary VerifyPromotedProductsTab(string tabName = "Promoted Products")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//li/span"), "'" + tabName + "' Tab Header Text not present.");
            Assert.AreEqual(tabName, driver._getText("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//li/span"), "'" + tabName + "' Header Text does not match.");

            //Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//div[@title='More Options']"), "'More Options' Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//div[@title='Export']"), "'Export' Icon not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id, '0chart')]//*[name()='svg']"), "'Bar Chart by " + tabName + "' not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]"), "'" + tabName + " and Values' Table not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[@class][1]"), "'" + tabName + "' Column Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[@class][2]"), "'Values' Column Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[not(@class)][1]"), "'" + tabName + "' Column not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[not(@class)][2]"), "'Values' Column not present.");

            if (!tabName.ToLower().Contains("subcategory") && !tabName.ToLower().Contains("retailer") && !tabName.ToLower().Contains("category") 
                && !tabName.ToLower().Contains("retailer group") && !tabName.ToLower().Contains("variety") && !tabName.ToLower().Contains("origin")
                && !tabName.ToLower().Contains("channel"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id, '1chart')]//*[name()='svg']"), "'Filtered Competition Chart for " + tabName + "' not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '1chartTabular')]"), "'Filtered Competition and Values' Table not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '1chartTabular')]//td[@class][1]"), "'Filtered Competition' Column Header not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '1chartTabular')]//td[@class][2]"), "'Values' Column Header not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '1chartTabular')]//td[not(@class)][1]"), "'Filtered Competition' Column not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '1chartTabular')]//td[not(@class)][2]"), "'Values' Column not present.");
            }

            Results.WriteStatus(test, "Pass", "Verified, " + tabName + " Tab");
            return new CategorySummary(driver, test);
        }
        
        ///<summary>
        ///Verify Ad Type/Medium by Retailers Tab
        ///</summary>
        ///<param name="tabName">Name of Tab to be Verified</param>
        ///<returns></returns>
        public CategorySummary VerifyAdType_MediumByRetailersTab(string tabName = "Ad Type")
        {

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//li/span"), "'" + tabName + "' Tab Header Text not present.");
            Assert.AreEqual(tabName +  " by Retailers", driver._getText("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//li/span"), "'" + tabName + "' Header Text does not match.");

            //Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//div[@title='More Options']"), "'More Options' Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id, 'Header')]//div[@title='Export']"), "'Export' Icon not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//div[contains(@id, '0chart')]//*[name()='svg']"), "'Bar Chart by " + tabName + "' not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]"), "'" + tabName + " and Values' Table not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[@class][1]"), "'" + tabName + "' Column Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[@class][2]"), "'A' Column Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[not(@class)][1]"), "'" + tabName + "' Column not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[not(@class)][2]"), "'A' Column not present.");

            if(tabName.ToLower().Contains("ad type"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[@class][3]"), "'B' Column Header not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[contains(@id, '0chartTabular')]//td[not(@class)][3]"), "'B' Column not present.");
            }

            Results.WriteStatus(test, "Pass", "Verified, " + tabName + " by Retailers Tab");
            return new CategorySummary(driver, test);
        }


        #endregion
    }
}
