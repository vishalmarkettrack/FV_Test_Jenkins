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
    public class PricingAndPromotions
    {
        #region Private Variables

        private IWebDriver pricingAndPromotions;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public PricingAndPromotions(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.pricingAndPromotions = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.pricingAndPromotions; }
            set { this.pricingAndPromotions = value; }
        }

        /// <summary>
        /// Verify Pricing Summary Screen
        /// </summary>
        /// <returns></returns>
        public PricingAndPromotions verifyPricingSummaryScreen()
        {
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/span");
            Assert.AreEqual(true, driver._getAttributeValue("xpath", "//div[@role='tabpanel']//li[@heading='Manufacturer Summary']", "class").Contains("active"));

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@ng-controller='PricingPromoByManController']//div[@class='divTabs vizTabs']//li[contains(@class,'ng-isolate')]"), "Different Tabs not present.");
            IList<IWebElement> tabCollections = driver._findElements("xpath", "//div[@ng-controller='PricingPromoByManController']//div[@class='divTabs vizTabs']//li[contains(@class,'ng-isolate')]");
            string[] tabNames = { "Manufacturer Summary", "Purchase Offer Types", "Pricing Profile" };

            for (int tab = 0; tab < tabCollections.Count; tab++)
                Assert.AreEqual(tabCollections[tab].Text, tabNames[tab], "'" + tabCollections[tab].Text + "' Tab Name not match with '" + tabNames[tab] + "'.");

            Results.WriteStatus(test, "Pass", "Verified Pricing & Promotions Screen.");
            return new PricingAndPromotions(driver, test);
        }

        #region Pricing Summary

        /// <summary>
        /// Verify Manufacture Summary tab in Detail
        /// </summary>
        /// <returns></returns>
        public PricingAndPromotions verifyManufacturerSummaryTab()
        {
            Assert.AreEqual(true, driver._isElementPresent("id", "dvCardContainerHeaderTitle"), "'Manufacturer Summary' Title not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='CustChartMainTbl']/div/div[@class='fa fa-question homeChartSearch']"), "Help icon not prensent.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='CustChartMainTbl']/div/div[@title='More Options']"), "More Options icon not prensent.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='CustChartMainTbl']/div/div[@title='Export']"), "Export icon not prensent.");
            Assert.IsTrue(driver._isElementPresent("id", "dvCardContainerView"), "'Manufacturer Summary Cart section not present.'");

            verifyPricingPromoCardInformation();
            return new PricingAndPromotions(driver, test);
        }

        /// <summary>
        /// Verify Pricing Promo Card Information in Detail
        /// </summary>
        /// <returns></returns>
        public PricingAndPromotions verifyPricingPromoCardInformation()
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvCardContainer']/div"), "Manufacturer Summary Card List not Present.");
            IList<IWebElement> cardCollections = driver._findElements("xpath", "//*[@id='dvCardContainer']/div");

            for (int i = 0; i < cardCollections.Count; i++)
            {
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]//div[contains(@class,'card_LabelText cursorPointer')]"), "'Manufacturer Name' not Present on Pricing Promo Card.");
                string manufacturerName = driver._getText("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]//div[contains(@class,'card_LabelText cursorPointer')]").Trim().Replace("\r\n", "");
                Assert.AreNotEqual("", manufacturerName, "'Manufacturer Name' not available on Pricing Promo Card.");

                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[@class='row color-white bg-card-blue']/div[2]/div/div[1]/div[1]"), "'TY # of Ads:' Label not Present for '" + manufacturerName + "' Manufacturer.");
                Assert.AreEqual("TY # of Ads:", driver._getText("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[@class='row color-white bg-card-blue']/div[2]/div/div[1]/div[1]").Trim().Replace("\r\n", ""));
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[@class='row color-white bg-card-blue']/div[2]/div/div[1]/div[2]"), "'TY # of Ads:' Number not Present for '" + manufacturerName + "' Manufacturer.");
                Assert.AreNotEqual("", driver._getText("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[@class='row color-white bg-card-blue']/div[2]/div/div[1]/div[2]"));

                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[@class='row color-white bg-card-blue']/div[2]/div/div[2]/div[1]"), "'% Chg YA :' Label not Present for '" + manufacturerName + "' Manufacturer.");
                Assert.AreEqual("% Chg YA :", driver._getText("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[@class='row color-white bg-card-blue']/div[2]/div/div[2]/div[1]").Trim().Replace("\r\n", ""), "'% Chg YA :' Label not match for '" + manufacturerName + "' Manufacturer.");
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[@class='row color-white bg-card-blue']/div[2]/div/div[2]/div[2]"), "'% Chg YA :' Percentage not Present for '" + manufacturerName + "' Manufacturer.");
                Assert.AreEqual(true, driver._getText("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[@class='row color-white bg-card-blue']/div[2]/div/div[2]/div[2]").Contains("%"), "'% Chg YA :' Percentage Label not match for '" + manufacturerName + "' Manufacturer.");

                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[2]/div[1]"), "'Min Unit Price - Sale' not Present for '" + manufacturerName + "' Manufacturer.");
                Assert.AreEqual(true, driver._getText("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[2]/div[1]").Trim().Replace("\r\n", "").Contains("MinUnit Price - Sale$"), "'Min Unit Price - Sale' Label not match for '" + manufacturerName + "' Manufacturer.");

                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[2]/div[2]"), "'Most Frequent Unit Price - Sale' not Present for '" + manufacturerName + "' Manufacturer.");
                Assert.AreEqual(true, driver._getText("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[2]/div[2]").Trim().Replace("\r\n", "").Contains("Most FrequentUnit Price - Sale$"), "'Most Frequent Unit Price - Sale' Label not match for '" + manufacturerName + "' Manufacturer.");

                Assert.AreEqual(true, driver._isElementPresent("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[2]/div[3]"), "'Average Unit Price - Sale' not Present for '" + manufacturerName + "' Manufacturer.");
                Assert.AreEqual(true, driver._getText("xpath", "//*[@id='dvCardContainer']/div[" + (i + 1) + "]/div/div[2]/div[3]").Trim().Replace("\r\n", "").Contains("AverageUnit Price - Sale$"), "'Average Unit Price - Sale' Label not match for '" + manufacturerName + "' Manufacturer.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Pricing Promo Card Information.");
            return new PricingAndPromotions(driver, test);
        }

        /// <summary>
        /// Verify Tooltip Text of Manufacture Help Icon
        /// </summary>
        /// <returns></returns>
        public PricingAndPromotions verifyTooltipTextOfHelpIcon()
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='CustChartMainTbl']/div/div[@class='fa fa-question homeChartSearch']"), "Help icon not prensent.");
            IWebElement helpIconEle = driver._findElement("xpath", "//div[@class='CustChartMainTbl']/div/div[@class='fa fa-question homeChartSearch']");
            driver.MouseHoverByJavaScript(helpIconEle);
            string msgHeader = "How are products priced compared to competitors? What is the most frequent promoted unit price?";
            string msgDetail = "Provides the minimum, most frequent and average price points promoted at a Manufacturer level. The color of the box is determined by the % change vs. year ago, if previous period is selected. Red = less, green = higher and yellow = same. (Gray indicates previous year value is not available.) Click into Manufacturer to see price trends by Retailer. Click on a data point and pull up product images.";
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvGrpLst' and contains(@style,'block')]//*[contains(text(),'" + msgHeader + "')]"), "'" + msgHeader + "' Tooltip header not match.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='dvGrpLst' and contains(@style,'block') and contains(text(),'" + msgDetail + "')]"), "'" + msgDetail + "' Tooltip In-Detail not match.");

            Results.WriteStatus(test, "Pass", "Verified Tooltip Text of Help Icon.");
            return new PricingAndPromotions(driver, test);
        }

        #endregion

        #endregion
    }
}
