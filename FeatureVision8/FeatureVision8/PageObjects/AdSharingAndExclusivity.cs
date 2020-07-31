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
    public class AdSharingAndExclusivity
    {
        #region Private Variables

        private IWebDriver adSharingAndExclusivity;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public AdSharingAndExclusivity(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.adSharingAndExclusivity = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.adSharingAndExclusivity; }
            set { this.adSharingAndExclusivity = value; }
        }

        ///<summary>
        ///Verify Ad Sharing And Exclusivity Page
        ///</summary>
        ///<returns></returns>
        public AdSharingAndExclusivity VerifyAdSharingAndExclusivityPage(string clientName = "Procter & Gamble", string searchName = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//h1[text()='Category & Brand Share']"), "Category & Brand Share Screen not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='title cursorpointer']/a[text()='Ad Sharing And Exclusivity']"), "Ad Sharing And Exclusivity link not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-media']/i[contains(@class, 'ad-sharing-and-exclusivity')]"), "Ad Sharing And Exclusivity icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-content']/div[contains(@class, 'submenudescription')]"), "Ad Sharing And Exclusivity description not present.");

            driver._click("xpath", "//div[@class='title cursorpointer']/a[text()='Ad Sharing And Exclusivity']");
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

            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Save Options']/i"), "'Save' option not present.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@role='tabpanel']//li[@heading]"), "Tabs not present.");
            string[] tabNameList = new string[] { "Exclusive and Shared Ad Blocks", "Shared Ad Blocks" };
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
            return new AdSharingAndExclusivity(driver, test);
        }

        ///<summary>
        ///Verify Exclusive and Shared Ad Blocks Tab
        ///</summary>
        ///<param name="viewBy">View By Dropdown Value</param>
        ///<returns></returns>
        public AdSharingAndExclusivity VerifyExclusiveAndSharedAdBlocksTab(string viewBy = "Brand")
        {
            if (!driver._getAttributeValue("xpath", "//div[@role='tabpanel']//li[@heading='Exclusive and Shared Ad Blocks']", "class").Contains("active"))
            {
                driver._click("xpath", "//div[@role='tabpanel']//li[@heading='Exclusive and Shared Ad Blocks']");
                Thread.Sleep(2000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@role='tabpanel']//li[@heading='Exclusive and Shared Ad Blocks']", "class").Contains("active"), "'Exclusive and Shared Ad Blocks' did not get selected.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvAdSharingAndExclusivityHeader']//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvAdSharingAndExclusivityHeader']//div[@title='More Options']"), "'More Options' Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvAdSharingAndExclusivityHeader']//div[@title='Export']"), "'Export' Icon not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@data-dropdown-index=0]//td[text()='View By:']"), "'View By' Text Label not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//button"), "'View By' Button not present.");

            if (viewBy.ToLower().Equals("brand"))
                viewBy = "Brand";
            else if (viewBy.ToLower().Equals("category"))
                viewBy = "Category";
            else if (viewBy.ToLower().Equals("manufacturer"))
                viewBy = "Manufacturer";

            driver._click("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//button");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//li"), "'View By' DDL not present");
            IList<IWebElement> ddlCollection = driver._findElements("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//li");

            bool avail = false;
            foreach (IWebElement element in ddlCollection)
                if (element.Text.ToLower().Contains(viewBy.ToLower()))
                {
                    avail = true;
                    element.Click();
                    break;
                }
            Assert.IsTrue(avail, "'" + viewBy + "' not found.");

            Thread.Sleep(2000);
            Assert.AreEqual(viewBy, driver._getAttributeValue("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//button", "title"), "'" + viewBy + "' View By Type not selected.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvAdSharingAndExclusivityHeaderTitle']//li"), "'Exclusive and Shared Ad Blocks' Heading not present.");
            Assert.AreEqual("Exclusive and Shared Ad Blocks by " + viewBy + " (Sort by Total Ad Boxes)", driver._getText("xpath", "//div[@id='dvAdSharingAndExclusivityHeaderTitle']//li/span"), "'Exclusive and Shared Ad Blocks' Heading text doesn't match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvAdSharingAndExclusivity0chart']//*[name()='svg']"), "'Column Chart by " + viewBy + "' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvAdSharingAndExclusivity0chart']//*[@class='highcharts-legend-item']//*[name()='tspan' and text()='Shared Ad Blocks']"), "'Shared Ad Blocks' legend not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvAdSharingAndExclusivity0chart']//*[@class='highcharts-legend-item']//*[name()='tspan' and text()='Exclusive Ad Blocks']"), "'Exclusive Ad Blocks' legend not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvAdSharingAndExclusivity0chartTabular_tblMain']"), "'Exclusive and Shared Ad Blocks' Table not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvAdSharingAndExclusivity0chartTabular_tblMain']//td[@class='HeaderColumn' and text()='Exclusive and Shared Ad Blocks by " + viewBy + " (Sort by Total Ad Boxes)']"), "Exclusive and Shared Ad Blocks by '" + viewBy + "' (Sort by Total Ad Boxes)' Table not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvAdSharingAndExclusivity0chartTabular_tblMain']//td[@class='HeaderColumn' and text()='Shared Ad Blocks']"), "Shared Ad Blocks' Table not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvAdSharingAndExclusivity0chartTabular_tblMain']//td[@class='HeaderColumn' and text()='Exclusive Ad Blocks']"), "Exclusive Ad Blocks' Table not present.");

            Results.WriteStatus(test, "Pass", "Verified, Exclusive and Shared Ad Blocks Tab");
            return new AdSharingAndExclusivity(driver, test);
        }

        ///<summary>
        ///Verify Ad Sharing And Exclusivity Help Icon
        ///</summary>
        ///<param name="tabName">Tab Name for Help Icon</param>
        ///<returns></returns>
        public AdSharingAndExclusivity VerifyAdSharingAndExclusivityPageHelpIcon(string tabName = "Exclusive And Shared Ad Blocks")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'active')]//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            driver._scrollintoViewElement("xpath", "//div[contains(@class,'active')]//div[@class='fa fa-question homeChartSearch']");
            driver.MouseHoverUsingElement("xpath", "//div[contains(@class,'active')]//div[@class='fa fa-question homeChartSearch']");

            string capturedHelpIconText = driver._getAttributeValue("xpath", "//div[contains(@class,'active')]//div[@onmouseout='HideTooltip();']", "onmouseover");
            string expectedHelpIconText1 = "Who are we sharing ad blocks with the most? How often do we have exclusive ad blocks versus competition? Are we sharing ad space with private label brands in the category?";
            string expectedHelpIconText2 = "Provides a total number of exclusive and shared ad blocks in a given period, filtered by brand, category or manufacturer. Drill down (click on any shard ad block bar in chart) to see what brand/category/manufacturer is being promoted with others.";

            if (!tabName.ToLower().Contains("exclusive"))
            {
                expectedHelpIconText1 = "What brands, manufacturers, and categories are being cross-promoted? Who is sharing ad blocks with who and how often?";
                expectedHelpIconText2 = "Provides a total number of shared ad blocks in a given period, filtered by brand, category or manufacturer.";
            }
            Console.WriteLine(capturedHelpIconText);
            Assert.IsTrue(capturedHelpIconText.ToLower().Contains(expectedHelpIconText1.ToLower()) && capturedHelpIconText.ToLower().Contains(expectedHelpIconText2.ToLower()), "Help Icon tooltip text did not match.");

            Results.WriteStatus(test, "Pass", "Verified, Help Icon for '" + tabName + "' Tab.");
            return new AdSharingAndExclusivity(driver, test);
        }

        ///<summary>
        ///Select Option From DropDown On Ad Sharing And Exclusivity Page
        ///</summary>
        ///<param name="menuIcon">Menu Icon to Click</param>
        ///<param name="optionName">Option to Click from DDL</param>
        ///<returns></returns>
        public AdSharingAndExclusivity SelectOptionFromDropDownOnAdSharingAndExclusivityPage(string menuIcon, string optionName)
        {
            string clickablePath = "";
            if (driver._isElementPresent("xpath", "//div[contains(@class,'active')]//div[@ng-show and @class='']//div[contains(@id, 'Header')]"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'active')]//div[@ng-show and @class='']//div[@title='" + menuIcon + "']"), "'" + menuIcon + "' Icon not present.");

                if (menuIcon.ToLower().Contains("window") || menuIcon.ToLower().Contains("more"))
                    menuIcon = "More Options";
                else
                    menuIcon = "Export";

                //driver._click("xpath", "//div[contains(@class,'active')]//div[@ng-show and @class='']//div[@title='" + menuIcon + "']");
                clickablePath = "//div[contains(@class,'active')]//div[@ng-show and @class='']//div[@title='" + menuIcon + "']";
                driver._click("xpath", clickablePath);
            }
            else if (driver._isElementPresent("xpath", "//div[contains(@class,'active')]//div[contains(@id, 'Header')]"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'active')]//div[@title='" + menuIcon + "']"), "'" + menuIcon + "' Icon not present.");

                if (menuIcon.ToLower().Contains("window") || menuIcon.ToLower().Contains("more"))
                    menuIcon = "More Options";
                else
                    menuIcon = "Export";

                //driver._click("xpath", "//div[contains(@class,'active')]//div[@title='" + menuIcon + "']");
                clickablePath = "//div[contains(@class,'active')]//div[@title='" + menuIcon + "']";
                driver._click("xpath", clickablePath);
            }
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@title='" + menuIcon + "']"), "'" + menuIcon + "' Icon not present.");

                if (menuIcon.ToLower().Contains("window") || menuIcon.ToLower().Contains("more"))
                    menuIcon = "More Options";
                else
                    menuIcon = "Export";

                //driver._click("xpath", "//div[@title='" + menuIcon + "']");
                clickablePath = "//div[@title='" + menuIcon + "']";
                driver._click("xpath", clickablePath);
            }

            Actions actions = new Actions(driver);
            if (driver._isElementPresent("xpath", "//ul[@data-headeroption-id and contains(@style, 'block')]//a") == false)
                driver._click("xpath", clickablePath);
            Assert.IsTrue(driver._isElementPresent("xpath", "//ul[@data-headeroption-id and contains(@style, 'block')]//a"), "DDL not present");
            actions.MoveToElement(driver.FindElement(By.XPath("//ul[@data-headeroption-id and contains(@style, 'block')]//li[1]//a"))).Perform();
            IList<IWebElement> ddlOptions = driver._findElements("xpath", "//ul[@data-headeroption-id and contains(@style, 'block')]//a");

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

            Thread.Sleep(10000);

            Results.WriteStatus(test, "Pass", "Selected, '" + optionName + "' Option From '" + menuIcon + "' DropDown On Ad Sharing And Exclusivity Page");
            return new AdSharingAndExclusivity(driver, test);
        }

        ///<summary>
        ///Verify Sort by Number Or Percentage On Exclusive And Shared Ad Blocks
        ///</summary>
        ///<param name="perc">Show by Percentage Or #</param>
        ///<returns></returns>
        public AdSharingAndExclusivity VerifySortByNumberOrPercentageOnExclusiveAndSharedAdBlocks(bool perc = false)
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

            return new AdSharingAndExclusivity(driver, test);
        }

        ///<summary>
        ///Verify Sort By Options On Exclusive And Shared Ad Blocks Tab
        ///</summary>
        ///<param name="sortBy">To Verify Sorting Type</param>
        ///<param name="viewBy">View By DDL Value</param>
        ///<returns></returns>
        public AdSharingAndExclusivity VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab(string viewBy = "Brand", string sortBy = "Total")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvAdSharingAndExclusivity0chartTabular_tblMain']//td[@class='HeaderColumn' and text()='Exclusive and Shared Ad Blocks by " + viewBy + " (Sort by " + sortBy + " Ad Boxes)']"), "Exclusive and Shared Ad Blocks by '" + viewBy + "' (Sort by " + sortBy + " Ad Boxes)' Table not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@class='MyCustGrid']//tr[not(@class)]/td[2]"), "Shared Ad Blocks Column not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@class='MyCustGrid']//tr[not(@class)]/td[3]"), "Exclusive Ad Blocks Column not present.");

            IList<IWebElement> sharedAdBlockColl = driver._findElements("xpath", "//table[@class='MyCustGrid']//tr[not(@class)]/td[2]");
            IList<IWebElement> exclusiveAdBlockColl = driver._findElements("xpath", "//table[@class='MyCustGrid']//tr[not(@class)]/td[3]");

            int[] sharedAdBlockList = new int[] { sharedAdBlockColl.Count };
            int[] exclusiveAdBlockList = new int[] { exclusiveAdBlockColl.Count };

            for (int i = 0; i < sharedAdBlockList.Length; i++)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", sharedAdBlockColl[i]);
                Assert.IsTrue(int.TryParse(sharedAdBlockColl[i].Text.Replace(",", ""), out sharedAdBlockList[i]), "Couldn't Convert '" + sharedAdBlockColl[i].Text + "' to int.");
                Assert.IsTrue(int.TryParse(exclusiveAdBlockColl[i].Text.Replace(",", ""), out exclusiveAdBlockList[i]), "Couldn't Convert '" + exclusiveAdBlockColl[i].Text + "' to int.");
            }

            int[] sortedArray = new int[] { sharedAdBlockList.Length };

            if (sortBy.ToLower().Contains("shared"))
            {
                Array.Copy(sharedAdBlockList, sortedArray, sharedAdBlockList.Length);
                Array.Sort(sortedArray);
                Array.Reverse(sortedArray);
                Assert.IsTrue(sharedAdBlockList.SequenceEqual(sortedArray), "'Sort by Shared' did not apply successfully.");
                Results.WriteStatus(test, "Pass", "'Sort by Shared' applied successfully.");
            }
            else if (sortBy.ToLower().Contains("exclusive"))
            {
                Array.Copy(exclusiveAdBlockList, sortedArray, sharedAdBlockList.Length);
                Array.Sort(sortedArray);
                Array.Reverse(sortedArray);
                Assert.IsTrue(exclusiveAdBlockList.SequenceEqual(sortedArray), "'Sort by Exclusive' did not apply successfully.");
                Results.WriteStatus(test, "Pass", "'Sort by Exclusive' applied successfully.");
            }
            else
            {
                int[] totalArray = new int[] { sharedAdBlockList.Length };
                for (int i = 0; i < totalArray.Length; i++)
                    totalArray[i] = sharedAdBlockList[i] + exclusiveAdBlockList[i];
                Array.Copy(totalArray, sortedArray, totalArray.Length);
                Array.Sort(sortedArray);
                Array.Reverse(sortedArray);
                Assert.IsTrue(totalArray.SequenceEqual(sortedArray), "'Sort by Total' did not apply successfully.");
                Results.WriteStatus(test, "Pass", "'Sort by Total' applied successfully.");
            }

            return new AdSharingAndExclusivity(driver, test);
        }

        ///<summary>
        ///Verify Shared Ad Blocks Tab
        ///</summary>
        ///<param name="viewBy">View By Dropdown Value</param>
        ///<returns></returns>
        public AdSharingAndExclusivity VerifySharedAdBlocksTab(string viewBy = "Brand")
        {
            if (!driver._getAttributeValue("xpath", "//div[@role='tabpanel']//li[@heading='Shared Ad Blocks']", "class").Contains("active"))
            {
                driver._click("xpath", "//div[@role='tabpanel']//li[@heading='Shared Ad Blocks']");
                Thread.Sleep(2000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@role='tabpanel']//li[@heading='Shared Ad Blocks']", "class").Contains("active"), "'Exclusive and Shared Ad Blocks' did not get selected.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMapHeader']//div[@class='fa fa-question homeChartSearch']"), "Help Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMapHeader']//div[@title='More Options']"), "'More Options' Icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMapHeader']//div[@title='Export']"), "'Export' Icon not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//td[text()='View By:']"), "'View By' Text Label not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//button"), "'View By' Button not present.");

            if (viewBy.ToLower().Equals("brand"))
                viewBy = "Brand";
            else if (viewBy.ToLower().Equals("category"))
                viewBy = "Category";
            else if (viewBy.ToLower().Equals("manufacturer"))
                viewBy = "Manufacturer";

            driver._click("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//button");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//li"), "'View By' DDL not present");
            IList<IWebElement> ddlCollection = driver._findElements("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//li");

            bool avail = false;
            foreach (IWebElement element in ddlCollection)
                if (element.Text.ToLower().Contains(viewBy.ToLower()))
                {
                    avail = true;
                    element.Click();
                    break;
                }
            Assert.IsTrue(avail, "'" + viewBy + "' not found.");

            Thread.Sleep(2000);
            Assert.AreEqual(viewBy, driver._getAttributeValue("xpath", "//div[contains(@class, 'active')]//table[@data-dropdown-index=0]//button", "title"), "'" + viewBy + "' View By Type not selected.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMapHeaderTitle']//li"), "'Exclusive and Shared Ad Blocks' Heading not present.");
            Assert.AreEqual("Shared Ad Blocks by " + viewBy + " (Sort by Total Ad Boxes)", driver._getText("xpath", "//div[@id='dvOccurenceOfHeatMapHeaderTitle']//li/span"), "'Shared Ad Blocks' Heading text doesn't match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[name()='svg']"), "'Column Chart by " + viewBy + "' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[@class='highcharts-legend-title']//*[name()='tspan' and text()='Ad Block Count']"), "'Ad Block Count']' legend not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[contains(@class,'tracker')]//*[name()='tspan']"), "'Ad Block Count' boxes not present");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[contains(@class,'tracker')]//*[name()='tspan' and not(text()=0)][1]"),"'Scrollable element not present.'");
            driver._scrollintoViewElement("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[contains(@class,'tracker')]//*[name()='tspan' and not(text()=0)][1]");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[contains(@class,'tracker')]//*[name()='tspan' and not(text()=0)][1]"), "Clickable Element not present.");
            int cnt = Convert.ToInt32(driver._getText("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[contains(@class,'tracker')]//*[name()='tspan' and not(text()=0)][1]"));
            driver._click("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[contains(@class,'tracker')]//*[name()='tspan' and not(text()=0)][1]"); Thread.Sleep(3000);

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='item ng-scope active']//li"), "'Carousels' not present.");
            if (cnt >= 5)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//span[@class='fa fa-angle-left icon']"), "Left Navigation Icon not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//span[@class='fa fa-angle-right icon']"), "Right Navigation Icon not present.");
            }
            Results.WriteStatus(test, "Pass", "Verified, Shared Ad Blocks Tab");
            return new AdSharingAndExclusivity(driver, test);
        }

        ///<summary>
        ///Verify Navigation On Shared Ad Blocks Tab
        ///</summary>
        ///<returns></returns>
        public AdSharingAndExclusivity VerifyNavigationOnSharedAdBlocksTab()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[contains(@class,'tracker')]//*[name()='tspan' and not(text()=0)]"), "'Ad Block Count' not present.");
            IList<IWebElement> adBlockCountColl = driver._findElements("xpath", "//div[@id='dvOccurenceOfHeatMap0chart']//*[contains(@class,'tracker')]//*[name()='tspan' and not(text()=0)]");

            foreach (IWebElement adBlockCount in adBlockCountColl)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", adBlockCount);
                string dateLinkText = adBlockCount.Text;
                int dateLinkNumber = 0;
                Assert.IsTrue(int.TryParse(dateLinkText, out dateLinkNumber), "Ad Block Count Text Could not be converted to int.");
                if (dateLinkNumber > 6)
                {
                    adBlockCount.Click();
                    break;
                }
            }

            Thread.Sleep(3000);
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='left carousel-control']"), "Left Navigation Button not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='right carousel-control']"), "Right Navigation Button not present.");

            string firstCard = driver._getAttributeValue("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//img", "ng-src");
            Actions action = new Actions(driver);
            action.Click(driver.FindElement(By.XPath("//span[@class='fa fa-angle-right icon']"))).Perform();
            Thread.Sleep(5000);
            string newCard = driver._getAttributeValue("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//img", "ng-src");
            Assert.AreNotEqual(firstCard, newCard, "Next Navigation button did not work.");

            action.Click(driver.FindElement(By.XPath("//span[@class='fa fa-angle-left icon']"))).Perform();
            Thread.Sleep(5000);
            string prevCard = driver._getAttributeValue("xpath", "//div[@class='carousel-inner']/div[contains(@class, 'active')]//li[1]//img", "ng-src");
            Assert.AreNotEqual(newCard, prevCard, "Previous Navigation button did not work.");
            Assert.AreEqual(firstCard, prevCard, "Previous Navigation button did not work.");

            Results.WriteStatus(test, "Pass", "Verified, Navigation On Shared Ad Blocks Carousel");
            return new AdSharingAndExclusivity(driver, test);
        }

        ///<summary>
        ///Verify Sort By Total Ad Blocks On Shared Ad Blocks Tab
        ///</summary>
        ///<param name="viewBy">View By DDL Option</param>
        ///<returns></returns>
        public AdSharingAndExclusivity VerifySortByTotalAdBlocksOnSharedAdBlocksTab(string viewBy = "Brand")
        {
            Assert.AreEqual("Shared Ad Blocks by " + viewBy + " (Sort by Total Ad Boxes)", driver._getText("xpath", "//div[@id='dvOccurenceOfHeatMapHeaderTitle']//li/span"), "'Shared Ad Blocks' Heading text doesn't match.");


            Results.WriteStatus(test, "Pass", "Verified, Sort By Total Ad Blocks On Shared Ad Blocks Tab");
            return new AdSharingAndExclusivity(driver, test);
        }

        ///<summary>
        ///Verify Sort By Shared Ad Blocks On Shared Ad Blocks Tab
        ///</summary>
        ///<param name="viewBy">View By DDL Option</param>
        ///<returns></returns>
        public AdSharingAndExclusivity VerifySortBySharedAdBlocksOnSharedAdBlocksTab(string viewBy = "Brand")
        {
            Assert.AreEqual("Shared Ad Blocks by " + viewBy + " (Sort by Shared Ad Boxes)", driver._getText("xpath", "//div[@id='dvOccurenceOfHeatMapHeaderTitle']//li/span"), "'Shared Ad Blocks' Heading text doesn't match.");

            Results.WriteStatus(test, "Pass", "Verified, Sort By Shared Ad Blocks On Shared Ad Blocks Tab");
            return new AdSharingAndExclusivity(driver, test);
        }





        #endregion
    }
}
