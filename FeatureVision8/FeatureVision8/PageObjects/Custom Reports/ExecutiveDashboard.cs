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
    public class ExecutiveDashboard
    {
        #region Private Variables

        private IWebDriver executiveDashboard;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public ExecutiveDashboard(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.executiveDashboard = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.executiveDashboard; }
            set { this.executiveDashboard = value; }
        }

        /// <summary>
        /// Verify Executive Dashboard Screen
        /// </summary>
        /// <returns></returns>
        public ExecutiveDashboard verifyExecutiveDashboardScreen()
        {
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/span");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='ContentWrapper']//div[@ng-controller='ExecutiveDashboardController']"), "'Executive Dashboard' Screen not load properly.");
            verifyTotalAlcoholBeverageAdBlocksSection();
            verifySalesAndAdShareSection();
            verifyKeySalesPerformanceMetricsSection();
            verifyAdFeaturePerformanceMetricsSection();

            Results.WriteStatus(test, "Pass", "Verified 'Executive Dashboard' Section.");
            return new ExecutiveDashboard(driver, test);
        }

        /// <summary>
        /// Verify Total Alcohol Beverage Ad Blocks Section
        /// </summary>
        /// <returns></returns>
        public ExecutiveDashboard verifyTotalAlcoholBeverageAdBlocksSection()
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvTotalAlkoholBeverageByAdBlocks']"), "'Total Alcohol Beverage Ad Blocks' Section not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvTotalAlkoholBeverageByAdBlocks']//div[@class='CustChartMainTbl']/div/div[@title='More Options']"), "More Options icon not prensent for 'Total Alcohol Beverage Ad Blocks' section.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvTotalAlkoholBeverageByAdBlocks']//div[@class='CustChartMainTbl']/div/div[@title='Export']"), "Export icon not prensent for 'Total Alcohol Beverage Ad Blocks' section.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvTotalAlkoholBeverageByAdBlocks0chart']"), "'Total Alcohol Beverage Ad Blocks' Chart not present.'");

            if (driver._isElementPresent("xpath", "//div[@id='dvTotalAlkoholBeverageByAdBlocks0chart']//p[@class='CustomMessageDisplay']"))
                Assert.IsTrue(driver._getText("xpath", "//div[@id='dvTotalAlkoholBeverageByAdBlocks0chart']//p[@class='CustomMessageDisplay']").Contains("Your search or filter parameters are not valid for this visualization."), "Message not match.");
            else
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvTotalAlkoholBeverageByAdBlocks0chart']//div[@class='highcharts-container']//*[name()='svg']"), "'Total Alcohol Beverage Ad Blocks' Chart not Display.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='radioButtonsParent']//*[text()='Ad Blocks']"), "'Ad Block' Radion option not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='radioButtonsParent']//*[text()='Brand Mentions']"), "'Brand Mentions' Radion option not present.");

            Results.WriteStatus(test, "Pass", "Verified 'Total Alcohol Beverage Ad Block' Section.");
            return new ExecutiveDashboard(driver, test);
        }

        /// <summary>
        /// Verify Sales & Ad Share Section
        /// </summary>
        /// <returns></returns>
        public ExecutiveDashboard verifySalesAndAdShareSection()
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvSalesAndAdShareByBrewer']"), "'Sales & Ad Share' Section not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvSalesAndAdShareByBrewer']//div[@class='CustChartMainTbl']/div/div[@title='More Options']"), "More Options icon not prensent for 'Total Alcohol Beverage Ad Blocks' section.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvSalesAndAdShareByBrewer']//div[@class='CustChartMainTbl']/div/div[@title='Export']"), "Export icon not prensent for 'Total Alcohol Beverage Ad Blocks' section.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvSalesAndAdShareByBrewer0chartP']"), "'$ Sales' Chart not present.'");
            if (driver._isElementPresent("xpath", "//div[@id='dvSalesAndAdShareByBrewer0chartP']//p[@class='CustomMessageDisplay']"))
                Assert.IsTrue(driver._getText("xpath", "//div[@id='dvSalesAndAdShareByBrewer0chartP']//p[@class='CustomMessageDisplay']").Contains("Your search or filter parameters are not valid for this visualization."), "Message not match.");
            else
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvSalesAndAdShareByBrewer0chartP']//div[@class='highcharts-container']/*[name()='svg']"), "'Sales' Chart not display.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvSalesAndAdShareByBrewer1chartP']"), "'Ad Features' Chart not present.'");
            if (driver._isElementPresent("xpath", "//div[@id='dvSalesAndAdShareByBrewer1chartP']//p[@class='CustomMessageDisplay']"))
                Assert.IsTrue(driver._getText("xpath", "//div[@id='dvSalesAndAdShareByBrewer1chartP']//p[@class='CustomMessageDisplay']").Contains("Your search or filter parameters are not valid for this visualization."), "Message not match.");
            else
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvSalesAndAdShareByBrewer1chartP']//div[@class='highcharts-container']/*[name()='svg']"), "'Ad Features' Chart not display.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='radioButtonsParent']//*[text()='Brewer']"), "'Brewer' Radion option not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='radioButtonsParent']//*[text()='Segment']"), "'Segment' Radion option not present.");

            Results.WriteStatus(test, "Pass", "Verified 'Sales & Ad Share' Section");
            return new ExecutiveDashboard(driver, test);
        }

        /// <summary>
        /// Verify Key Saled Preformance Metrics Section
        /// </summary>
        /// <returns></returns>
        public ExecutiveDashboard verifyKeySalesPerformanceMetricsSection()
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvKeySalesPerformanceGridHeaderText']"), "'Key Sales Performance Metrics' Section not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvKeySalesPerformanceGridHeaderText']//div[@class='CustChartMainTbl']/div/div[@title='Export']"), "Export icon not prensent for 'Key Sales Performance Metrics' section.");

            if (driver._isElementPresent("xpath", "//div[@id='dvKeySalesPerformanceGrid']/p[@class='CustomMessageDisplay']"))
                Assert.IsTrue(driver._getText("xpath", "//div[@id='dvKeySalesPerformanceGrid']//p[@class='CustomMessageDisplay']").Contains("Your search or filter parameters are not valid for this visualization."), "Message not match.");
            else
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvKeySalesPerformanceGrid_tblMain']"), "'Key Sales Performance Metrics' Grid not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvKeySalesPerformanceGrid']"), "'Key Sales Performance Metrics' Grid not present.");

            Results.WriteStatus(test, "Pass", "Verified 'Key Sales Performance Metrics' Section");
            return new ExecutiveDashboard(driver, test);
        }

        /// <summary>
        /// Verify Ad Feature Performance Metrics Section
        /// </summary>
        /// <returns></returns>
        public ExecutiveDashboard verifyAdFeaturePerformanceMetricsSection()
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvAdFeaturePerformanceGridHeaderText']"), "'Ad Feature Performance Metrics' Section not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvAdFeaturePerformanceGridHeaderText']//div[@class='CustChartMainTbl']/div/div[@title='Export']"), "Export icon not prensent for 'Ad Feature Performance Metrics' section.");

            if (driver._isElementPresent("xpath", "//div[@id='dvAdFeaturePerformanceGrid']/p[@class='CustomMessageDisplay']"))
                Assert.IsTrue(driver._getText("xpath", "//div[@id='dvAdFeaturePerformanceGrid']//p[@class='CustomMessageDisplay']").Contains("Your search or filter parameters are not valid for this visualization."), "Message not match.");
            else
                Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='dvAdFeaturePerformanceGrid_tblMain']"), "'Ad Feature Performance Metrics' Grid not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvAdFeaturePerformanceGrid']"), "'Ad Feature Performance Metrics' Grid not present.");

            Results.WriteStatus(test, "Pass", "Verified 'Ad Feature Performance Metrics' Section");
            return new ExecutiveDashboard(driver, test);
        }


        public ExecutiveDashboard clickButtonAndSelectOptionFromList(string sectionName, string menuIcon, string option)
        {
            string sectionID = "dvTotalAlkoholBeverageByAdBlocksHeader";
            if (sectionName == "Sales & Ad Share")
                sectionID = "dvSalesAndAdShareByBrewerHeader";
            if (sectionName == "Key Sales Performance Metrics")
                sectionID = "dvKeySalesPerformanceGridHeader";
            if (sectionName == "Ad Feature Performance Metrics")
                sectionID = "dvAdFeaturePerformanceGridHeader";

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='" + sectionID + "']"), "'" + sectionName + "' Section not Present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='" + sectionID + "']//div[@title='" + menuIcon + "']"), "'" + menuIcon + "' Icon not Present for '" + sectionName + "' Section.");
            driver._clickByJavaScriptExecutor("//div[@id='" + sectionID + "']//div[@title='" + menuIcon + "']");
            driver.MouseHoverUsingElement("xpath", "//ul[@data-headeroption-id and contains(@style, 'block')]//li[1]");

            Actions actions = new Actions(driver);
            if (driver._isElementPresent("xpath", "//ul[@data-headeroption-id and contains(@style, 'block')]//a") == false)
                driver._click("xpath", "//div[@id='" + sectionID + "']//div[@title='" + menuIcon + "']");
            Assert.IsTrue(driver._isElementPresent("xpath", "//ul[@data-headeroption-id and contains(@style, 'block')]//a"), "DDL not present");
            actions.MoveToElement(driver.FindElement(By.XPath("//ul[@data-headeroption-id and contains(@style, 'block')]//li[1]//a"))).Perform();
            IList<IWebElement> ddlOptions = driver._findElements("xpath", "//ul[@data-headeroption-id and contains(@style, 'block')]//a");

            bool avail = false;
            foreach (IWebElement optionList in ddlOptions)
            {
                actions.MoveToElement(optionList).Perform();
                if (optionList.Text.ToLower().Equals(option.ToLower()))
                {
                    avail = true;
                    optionList.Click();
                    break;
                }
            }
            Assert.IsTrue(avail, "'" + option + "' not found."); Thread.Sleep(10000);

            Results.WriteStatus(test, "Pass", "Selected, '" + option + "' Option From '" + menuIcon + "' DropDown On '" + sectionName + "' Section.");
            return new ExecutiveDashboard(driver, test);
        }

        public string VerifyFileDownloadedOrNotOnScreen(string fileName, string FileType)
        {
            bool Exist = false;
            string FilePath = "";
            string Path = ExtentManager.ResultsDir;
            string[] filePaths = Directory.GetFiles(Path, FileType);

            foreach (string filePath in filePaths)
            {
                FileInfo ThisFile = new FileInfo(filePath);
                if (filePath.Contains(fileName + "-" + DateTime.Today.ToString("MM-DD-yyyy")) || filePath.Contains(fileName))
                {
                    ThisFile = new FileInfo(filePath);
                    FilePath = filePath;
                    if (ThisFile.LastWriteTime.ToShortTimeString() == DateTime.Now.ToShortTimeString() ||
                        ThisFile.LastWriteTime.AddMinutes(1).ToShortTimeString() == DateTime.Now.ToShortTimeString() ||
                        ThisFile.LastWriteTime.AddMinutes(2).ToShortTimeString() == DateTime.Now.ToShortTimeString() ||
                        ThisFile.LastWriteTime.AddMinutes(3).ToShortTimeString() == DateTime.Now.ToShortTimeString() ||
                        ThisFile.LastWriteTime.AddMinutes(4).ToShortTimeString() == DateTime.Now.ToShortTimeString())
                    {
                        Exist = true;
                        break;
                    }
                }
            }

            Assert.AreEqual(true, Exist, "'" + fileName + "' " + FileType.Replace("*", "").ToUpper() + "' File Not Exported Properly.");
            Results.WriteStatus(test, "Pass", "Verified, <b>'" + FileType.Replace("*", "").ToUpper() + "'</b> File Exported Properly for '" + fileName + "' Report File.");
            return FilePath;
        }


        #endregion
    }
}
