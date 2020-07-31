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
    public class SummarySection
    {
        #region Private Variables

        private IWebDriver summarySection;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public SummarySection(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.summarySection = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.summarySection; }
            set { this.summarySection = value; }
        }

        /// <summary>
        /// Verify Summary Gris Section
        /// </summary>
        /// <returns></returns>
        public SummarySection verifySummaryGridSection()
        {
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvSummaryReportHeaderTitle']"), "Summary Header not Present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='summaryReportHeaderLinkParent']/a[@title='Edit Summary Options']"), "'Edit' Summary Options link not Present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='summaryReportHeaderLinkParent']/a[@title='View Saved Summary Templates']"), "'Saved View' Summary Templates link not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvSummaryGrid']"), "Summary Grid not Present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//ul[contains(@class, 'pagination') and contains(@ng-change,'summaryGridPage')]/li/a"), "Pagination not present.");

            Results.WriteStatus(test, "Pass", "Verified, Summary Grid Section.");
            return new SummarySection(driver, test);
        }

        ///// <summary>
        ///// Click Plus Icon and Verify Summary Options popup
        ///// </summary>
        ///// <returns></returns>
        //public SummarySection clickPlusIconAndVerifySummaryOptionsPopup(string defaultTab = "Define Summary")
        //{
        //    Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu[@data='menuOptions']//a[@title='Create Summary']"), "'Create Summary' (Plus icon) not present.");
        //    driver._clickByJavaScriptExecutor("//navigation-menu[@data='menuOptions']//a[@title='Create Summary']");
        //    driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
        //    driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");

        //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//h4[text()='Summary Options']"), "'Summary Options' Popup Window with Header not present.");
        //    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div[@id='divTabs']/div[@class='ng-isolate-scope']/ul/li"), "Summary Options tabs not present.");
        //    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='divTabs']/div[@class='ng-isolate-scope']/ul/li[contains(@class,'active')]/a[text()='"+ defaultTab + "']"), "'" + defaultTab + "' Tab Default not opened.");

        //    Results.WriteStatus(test, "Pass", "Clicked, Plus icon and Verified 'Summary Options' popup window.");
        //    return new SummarySection(driver, test);
        //}

        /// <summary>
        /// Click Edit Summary Option link
        /// </summary>
        /// <returns></returns>
        public SummarySection clickEditLinkFromSummaryGridSection(string defaultTab = "Define Summary")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='summaryReportHeaderLinkParent']/a[@title='Edit Summary Options']"), "'Edit' Summary Options link not Present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='summaryReportHeaderLinkParent']/a[@title='Edit Summary Options' and text()='Edit']"), "'Edit' Text not match for 'Edit Summary Option'.");
            driver._clickByJavaScriptExecutor("//div[@class='summaryReportHeaderLinkParent']/a[@title='Edit Summary Options']");
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//h4[text()='Summary Options']"), "'Summary Options' Popup Window with Header not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div[@id='divTabs']/div[@class='ng-isolate-scope']/ul/li"), "Summary Options tabs not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='divTabs']/div[@class='ng-isolate-scope']/ul/li[contains(@class,'active')]/a[text()='" + defaultTab + "']"), "'" + defaultTab + "' Tab Default not opened.");

            Results.WriteStatus(test, "Pass", "Clicked, Plus icon and Verified 'Summary Options' popup window.");
            return new SummarySection(driver, test);
        }

        /// <summary>
        /// Click Tab from Summary Options popup window
        /// </summary>
        /// <param name="tabName">Tab Name</param>
        /// <returns></returns>
        public SummarySection clickTabFromSummaryOptionsPopupWindow(string tabName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div[@id='divTabs']/div[@class='ng-isolate-scope']/ul/li"), "Summary Options tabs not present.");
            IList<IWebElement> tabs = driver.FindElements(By.XPath("//div[@class='modal-content']//div[@id='divTabs']/div[@class='ng-isolate-scope']/ul/li"));
            bool avail = false;
            for (int i = 0; i < tabs.Count; i++)
            {
                if (tabs[i].Text.Contains(tabName))
                {
                    tabs[i].Click(); Thread.Sleep(3000);
                    driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
                    Assert.IsTrue(tabs[i].GetAttribute("class").Contains("active"), "'" + tabName + "' Tab not Active.");
                    avail = true;
                    break;
                }
            }

            Assert.IsTrue(avail, "'" + tabName + "' Tab not present on Summary Options popup window.");
            Results.WriteStatus(test, "Pass", "Clicked '" + tabName + "' Tab from Summary Options popup window.");
            return new SummarySection(driver, test);
        }

        #region Numeric Summary

        /// <summary>
        /// Verify Numeric Summary Section
        /// </summary>
        /// <returns></returns>
        public SummarySection verifyNumericSummarySection()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div[contains(@class,'active')]/div[contains(@src,'numericSummaryTab')]"), "Numeric Summary Section not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[contains(text(),'Select Breakouts:')]"), "'Select Breakouts:' label not present or match.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[1]//div[@class='dropdown btn-group']/button"), "Select Format Dropdown list not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[contains(text(),'Select the numeric field and highest/lowest values to be used in your report:')]"), "'Select the numeric field and highest/lowest values to be used in your report:' label not present or match.");

            string[] gridLabels = { "Base The Report On This Field:", "The Lowest Value In Your Selected Data is:", "The Highest Value In Your Selected Data is:", "Use This As The Lowest Value For The Report:", "Use This As The Highest Value For The Report:" };
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[3]//div[@class='row']"), "Grid Section not present.");
            IList<IWebElement> labels = driver.FindElements(By.XPath("//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[3]//div[@class='row']//div[contains(@class,'cell-lebel')]"));
            bool avail = false;

            for (int grid = 0; grid < gridLabels.Length; grid++)
            {
                for (int i = 0; i < labels.Count; i++)
                {
                    if (labels[i].Text.Contains(gridLabels[grid]))
                    {
                        avail = true;
                        break;
                    }
                }
                Assert.IsTrue(avail, "'" + gridLabels[grid] + "' Label not Present on Grid section.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[3]//div[@class='row'][2]"), "Value Section not present of Grid.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[3]//div[@class='row'][2]//button[@role='button']"), "Dropdown not present for 'Base The Report On This Field:'.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[3]//div[@class='row'][2]/div/div[2]"), "Value for 'The Lowest Value In Your Selected Data is:' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[3]//div[@class='row'][2]/div/div[3]"), "Value for 'The Highest  Value In Your Selected Data is:' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[3]//div[@class='row'][2]/div/div[4]//input"), "Input Text for 'Use This As The Lowest Value For The Report:' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[3]//div[@class='row'][2]/div/div[5]//input"), "Input Text for 'Use This As The Highest Value For The Report:' not present.");

            string[] buttonsName = { "Run Report", "Save", "Clear", "Cancel" };
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']/div[@class='modal-footer']//button[@type='button']"), "Buttons not present on Popup window.");
            //IList<IWebElement> buttonCollections = driver.FindElements(By.XPath("//div[@class='modal-content']/div[@class='modal-footer']//button[@type='button']"));

            for (int i = 0; i < buttonsName.Length; i++)
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']/div[@class='modal-footer']//button[@type='button' and text()='" + buttonsName[i] + "']"), "'" + buttonsName[i] + "' Button not present.");

            Results.WriteStatus(test, "Pass", "Verified Numeric Summary Section in Detail.");
            return new SummarySection(driver, test);
        }

        /// <summary>
        /// click Select Format dropdown and Select Option
        /// </summary>
        /// <param name="optionName">Option Name</param>
        /// <returns></returns>
        public SummarySection clickSelectFormatDropdownAndSelectOptionFromList(string optionName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[1]//div[@class='dropdown btn-group']/button"), "Select Format Dropdown list not present.");

            IWebElement dropdpwnSelectionName = driver.FindElement(By.XPath("//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[1]//div[@class='dropdown btn-group']/button"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", dropdpwnSelectionName);
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[1]//div[@class='dropdown btn-group open']"), "Select Format Dropdown list not opened.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[1]//div[@class='dropdown btn-group open']/ul/li"), "Selcet Format Dropdown list not display.");
            IList<IWebElement> collections = driver.FindElements(By.XPath("//div[contains(@src,'numericSummaryTab')]//div[@class='customScrollBar']/div[1]//div[@class='dropdown btn-group open']/ul/li"));

            bool avail = false;
            for (int i = 0; i < collections.Count; i++)
            {
                if (collections[i].Text == optionName)
                {
                    collections[i].Click(); Thread.Sleep(1000);
                    Assert.IsTrue(dropdpwnSelectionName.Text.Contains(optionName), "'" + optionName + "' Option not Selected.");
                    avail = true;
                    break;
                }
            }

            Assert.IsTrue(avail, "'" + optionName + "' Option not present on Dropdown list.");
            Results.WriteStatus(test, "Pass", "clicked, Format Options drodpown and Selected <b>'" + optionName + "'</b> Option from list.");
            return new SummarySection(driver, test);
        }

        /// <summary>
        /// click Button from Summary Options popup window
        /// </summary>
        /// <param name="buttonName">Button Name</param>
        /// <returns></returns>
        public SummarySection clickButtonFromSummaryOptionsPopupWindow(string buttonName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']/div[@class='modal-footer']//button[@type='button' and text()='" + buttonName + "']"), "'" + buttonName + "' Button not present.");
            driver._clickByJavaScriptExecutor("//div[@class='modal-content']/div[@class='modal-footer']//button[@type='button' and text()='" + buttonName + "']");

            Results.WriteStatus(test, "Pass", "Clicked '" + buttonName + "' Button from Summary Options Popup Window.");
            return new SummarySection(driver, test);
        }

        #endregion


        #region Old Numeric Summary

        public SummarySection verifySummaryOptionsAndClickTab(string tabName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='popupDiv1']//*[text()='Summary Options']"), "'Summary Options' popup window not prenset.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//ul[@id='ulSummTabs']/li"), "Tabs not present.");
            IList<IWebElement> tabCollections = driver.FindElements(By.XPath("//ul[@id='ulSummTabs']/li"));

            driver._clickByJavaScriptExecutor("//ul[@id='ulSummTabs']/li");

            return new SummarySection(driver, test);
        }

        public SummarySection verifySummaryOptionsPopupWindow(string optionName, string activeTabName = "Define Summary")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='popupDiv1']//*[text()='Summary Options']"), "'Summary Options' popup window not prenset.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//ul[@id='ulSummTabs']/li"), "Tabs not present.");
            IList<IWebElement> tabCollections = driver.FindElements(By.XPath("//ul[@id='ulSummTabs']/li"));

            Assert.IsTrue(driver._isElementPresent("xpath", "//ul[@id='ulSummTabs']/li[contains(@class,'active')]//*[text()='" + activeTabName + "']"), "'" + activeTabName + "' Tab Default not active.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvNumReportComb']/button"), "Select Format Drodpdown not present.");
            driver._click("xpath", "//div[@id='dvNumReportComb']/button");
            driver.MouseHoverUsingElement("xpath", "//table[@id='tbldvNumReportComb']//tr[1]");

            IList<IWebElement> optionsCollections = driver.FindElements(By.XPath("//table[@id='tbldvNumReportComb']//tr"));
            for (int j = 0; j < optionsCollections.Count; j++)
                Console.WriteLine(optionsCollections[j].Text);

            for (int i = 0; i < optionsCollections.Count; i++)
            {
                if (optionsCollections[i].Text == optionName)
                {
                    optionsCollections[i].Click(); Thread.Sleep(2000);
                    break;
                }
            }
            Results.WriteStatus(test, "Pass", "clicked, Format Options drodpown and Selected <b>'" + optionName + "'</b> Option from list.");

            return new SummarySection(driver, test);
        }

        public SummarySection clickDisplayReportButtonOnSummaryOptionsSection()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='SumOptDisplay']"), "'Display Report' Button not present.");
            driver._clickByJavaScriptExecutor("//input[@id='SumOptDisplay']");
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");

            Results.WriteStatus(test, "Pass", "Clicked 'Display Report' Button from Summary Options Popup Window.");

            return new SummarySection(driver, test);
        }

        public SummarySection clickPlusIcon()
        {
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");
            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu[@data='menuOptions']//a[@title='Create Summary']"), "'Create Summary' (Plus icon) not present.");
            driver._clickByJavaScriptExecutor("//navigation-menu[@data='menuOptions']//a[@title='Create Summary']");
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");

            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='popupDiv1']//*[text()='Summary Options']"), "'Summary Options' popup window not prenset.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//ul[@id='ulSummTabs']/li"), "Tabs not present.");

            Results.WriteStatus(test, "Pass", "Clicked, Plus icon and Verified 'Summary Options' popup window.");
            return new SummarySection(driver, test);
        }

        #endregion

        #endregion
    }
}
