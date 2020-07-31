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
    public class MyReports
    {
        #region Private Variables

        private IWebDriver myReports;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public MyReports(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.myReports = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.myReports; }
            set { this.myReports = value; }
        }

        /// <summary>
        /// Select and Verify tab from My Reports section
        /// </summary>
        /// <param name="tabName">Tab Name to select</param>
        /// <returns></returns>
        public MyReports selectAndVerifyTabFromMyReportsSection(string tabName)
        {
            IList<IWebElement> tabsList = driver.FindElements(By.XPath("//div[@role='tabpanel']/div[@class='ng-isolate-scope']/ul/li"));
            string[] tabTitles = { "Saved Results", "Subscriptions", "Exported Reports", "Scorecard Chart", "Best Practice Reports" };
            bool avail = false;

            for (int i = 0; i < tabTitles.Count(); i++)
                if (tabsList[i].Text == tabName)
                {
                    driver._clickByJavaScriptExecutor("//div[@role='tabpanel']/div[@class='ng-isolate-scope']/ul/li[" + (i + 1) + "]/a"); Thread.Sleep(3000);
                    driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
                    driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");
                    Assert.AreEqual(true, tabsList[i].GetAttribute("class").Contains("active"), "'" + tabName + "' tab not selected.");
                    avail = true;
                    break;
                }

            Assert.AreEqual(true, avail, "'" + tabName + "' tab not Present.");
            Results.WriteStatus(test, "Pass", "Selected and Verified '" + tabName + "' tab from My Reports.");
            return new MyReports(driver, test);
        }

        #region Saved Results

        /// <summary>
        /// Verify Saved Results Screen
        /// </summary>
        /// <param name="noRecords">Check No Records message</param>
        /// <param name="inDetails">Verify tab in Detail</param>
        /// <returns></returns>
        public MyReports verifySavedResultsScreen(bool noRecords = false, bool inDetails = true)
        {
            if (driver._getAttributeValue("xpath", "//div[@data-loader]//div[@class='ProcessLoader']", "style") == "")
                driver._waitForElementToBePopulated("xpath", "//div[@data-loader and contains(@style,'display')]//div[@class='ProcessLoader']");

            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='text-box-input']"), "Search Area not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='text-box-input' and @placeholder='Search']"), "'Search' watermark not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@role='tabpanel']/div[@class='ng-isolate-scope']/ul/li"), "Tabs not present.");

            IList<IWebElement> tabsList = driver.FindElements(By.XPath("//div[@role='tabpanel']/div[@class='ng-isolate-scope']/ul/li"));
            //string[] tabTitles = { "Saved Results", "Subscriptions", "Exported Reports", "Scorecard Chart", "Best Practice Reports" };
            string[] tabTitles = { "Saved Results", "Subscriptions", "Exported Reports", "Scorecard Chart", "Best Practice Rpt" };

            for (int i = 0; i < tabTitles.Count(); i++)
                Assert.AreEqual(tabTitles[i], tabsList[i].Text, "'" + tabTitles[i] + "' tab not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//button[contains(@ng-click, 'showSearchListDropdown')]"), "Show dropdown not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//button[contains(@ng-click, 'createdSearchListDropdown')]"), "Created By dropdown not present.");

            if (inDetails == true)
                if (noRecords)
                    Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[contains(@ng-style,'myRptCntrl')]//*[contains(text(),'No records found.')]"), "'No records found.' message not present.");
                else
                {
                    IList<IWebElement> columnHeaderColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@id='borderLayout_eRootPanel']//div[@class='ag-header-container']//div[@class='ag-header-row']/div");
                    string[] columnHeaderNamesList = new string[] { "Search Name", "Created By", "Label", "Type", "Last Run", "Created" };

                    foreach (string columnHeaderName in columnHeaderNamesList)
                    {
                        bool avail = false;
                        foreach (IWebElement columnHeader in columnHeaderColl)
                            if (columnHeader.Text.ToLower().Equals(columnHeaderName.ToLower()))
                            {
                                avail = true;
                                break;
                            }
                        Assert.IsTrue(avail, "'" + columnHeaderName + "' column header name not found.");
                    }
                }

            Results.WriteStatus(test, "Pass", "Verified, Saved Results scection.");
            return new MyReports(driver, test);
        }

        ///<summary>
        ///Verify Created By Dropdown
        ///</summary>
        ///<param name="option">Option to select from Show Dropdown</param>
        ///<returns></returns>
        public MyReports VerifyCreatedByDropdown(string option)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[text()='Created By:']"), "Show dropdown label not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//button[contains(@ng-click, 'createdSearchListDropdown')]"), "Created By dropdown not present.");

            driver._clickByJavaScriptExecutor("//button[contains(@ng-click, 'createdSearchListDropdown')]");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@ng-class, 'createdSearchListDropdown')]//li"), "Show DDL not present.");
            IList<IWebElement> showDDLColl = driver._findElements("xpath", "//div[contains(@ng-class, 'createdSearchListDropdown')]//li");

            string[] ddlLists = { "All", "Me", "Client", "Numerator" };
            for (int j = 0; j < showDDLColl.Count(); j++)
                Assert.AreEqual(true, ddlLists[j].Contains(showDDLColl[j].Text), "'" + ddlLists[j] + "' tab not present.");

            foreach (IWebElement showEle in showDDLColl)
                if (showEle.Text.ToLower().Contains(option.ToLower()))
                {
                    showEle.Click();
                    break;
                }

            Assert.AreEqual(true, driver._getText("xpath", "//button[contains(@ng-click, 'createdSearchListDropdown')]").ToLower().Contains(option.ToLower()), "'" + option + "' not selected");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SharedType']"), "Cells not present in Query type column.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SharedType']");

            for (int i = 0; i < cellCollection.Count; i++)
            {
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + i + "]/div[@colid='SharedType']"), "Created By not Present from list.");
                driver._scrollintoViewElement("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + i + "]/div[@colid='SharedType']");
                Assert.AreEqual(option.ToLower(), cellCollection[i].Text.ToLower(), "Created By '" + option + "' option was not applied successfully.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Created By Dropdown");
            return new MyReports(driver, test);
        }

        /// <summary>
        /// Get Saved Results name from list
        /// </summary>
        /// <returns></returns>
        public String getSavedResultNameFromList(string tabName = "Saved Results")
        {
            string savedResultName = ""; string tabColId = "SavedAs";
            if (tabName == "Subscriptions")
                tabColId = "SubscriptionName";
            if (tabName == "Exported Reports" || tabName == "Best Practice Reports")
                tabColId = "ReportName";

            if (driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@class='ag-body-container']//div[@row]//div[@colid='" + tabColId + "']"))
                savedResultName = driver._getText("xpath", "//div[@class='tab-pane ng-scope active']//div[@class='ag-body-container']//div[@row]//div[@colid='" + tabColId + "']");
            else
                Assert.IsFalse(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@class='ag-body-container']//div[@row]//div[@colid='" + tabColId + "']"), "Saved Searches list not present.");

            Results.WriteStatus(test, "Pass", "Verified " + tabName + " list and get " + tabName + " Name : " + savedResultName + ".");
            return savedResultName;
        }

        ///<summary>
        /// Insert Value in Search box and Verify with Saved Name List
        ///</summary>
        ///<param name="searchName">Inserted Search Value</param>
        ///<returns></returns>
        public MyReports insertValueInSearchBoxAndVerifyWithSavedNameList(string searchName, string tabName = "Saved Results")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='text-box-input']"), "Search Box not present.");
            Assert.AreEqual("Search", driver._getAttributeValue("xpath", "//input[@id='text-box-input']", "placeholder"), "'Search Box' placeholder text does not match.");

            driver._type("xpath", "//input[@id='text-box-input']", searchName);
            Thread.Sleep(1000);

            string tabColId = "SavedAs";
            if (tabName == "Subscriptions")
                tabColId = "SubscriptionName";
            if (tabName == "Exported Reports" || tabName == "Best Practice Reports")
                tabColId = "ReportName";

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='" + tabColId + "']"), "Cells not present in Query type column.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='" + tabColId + "']");

            for (int i = 0; i < cellCollection.Count; i++)
            {
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + i + "]/div[@colid='" + tabColId + "']"), "'Saved Results' Name not present from List.");
                driver._scrollintoViewElement("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + i + "]/div[@colid='" + tabColId + "']");
                Assert.IsTrue(cellCollection[i].Text.ToLower().Contains(searchName.ToLower()), "'" + searchName + "' Saved Search was not applied successfully.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Inserted Search Name : " + searchName + " with " + tabName + " Saved List.");
            return new MyReports(driver, test);
        }

        ///<summary>
        ///Verify Saved Search Options
        ///</summary>
        ///<param name="option">Option to click</param>
        ///<returns></returns>
        public string VerifySavedResultsOptions(string option = "", bool createdByMe = false)
        {
            string searchName = "";
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid=0]//div[@class]"), "Saved Search Options icons not present");
            IList<IWebElement> savedSearchOptionsIconColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid=0]//div[@class]");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SavedAs']"), "Saved Search Options icons not present");
            IList<IWebElement> savedSearchNameColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SavedAs']");

            Random rand = new Random();
            int x = rand.Next(0, savedSearchOptionsIconColl.Count);

            driver._scrollintoViewElement("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + x + "]/div[@colid=0]");
            searchName = savedSearchNameColl[x].Text;

            driver._clickByJavaScriptExecutor("//div[@class='tab-pane ng-scope active']//div[@row='" + x + "']/div[@colid=0]//div[@class]");
            //savedSearchOptionsIconColl[x].Click();
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + x + "]/div[@colid='0']//li"), "Saved Search option DDL not present.");
            IList<IWebElement> optionsDDLColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + x + "]/div[@colid='0']//li");
            string[] optionsNameList = new string[] { "Run", "Manage Label" };
            if (createdByMe)
            {
                Array.Resize(ref optionsNameList, 3);
                optionsNameList[1] = "Delete";
                optionsNameList[2] = "Manage Label";
            }

            IWebElement optionButton = null;

            foreach (string optionName in optionsNameList)
            {
                bool avail = false;
                foreach (IWebElement optionDDLEle in optionsDDLColl)
                    if (optionDDLEle.Text.ToLower().Contains(optionName.ToLower()))
                    {
                        avail = true;
                        if (optionName.ToLower().Equals(option.ToLower()))
                            optionButton = optionDDLEle;
                        break;
                    }
                Assert.IsTrue(avail, "'" + optionName + "' not found in Left Navigation Menu.");
            }

            if (option != "")
            {
                Assert.AreNotEqual(null, optionButton, "'" + option + "' button not present.");
                optionButton.Click();
                Thread.Sleep(1000);
                Results.WriteStatus(test, "Pass", "Clicked, '" + option + "' option from Saved Search Options");
            }

            Results.WriteStatus(test, "Pass", "Verified, Saved Search Options");
            return searchName;
        }

        ///<summary>
        ///Verify Drag And Drop Functionality On Saved Search Columns
        ///</summary>
        ///<returns></returns>
        public MyReports VerifyDragAndDropFunctionalityOnSavedResultsColumns(string columnOne, string columnTwo)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[contains(@class, 'ag-header-cell-sortable')]//span[@id='agText']"), "Column Headers not presentin table.");
            IList<IWebElement> columnHeaderColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[contains(@class, 'ag-header-cell-sortable')]//span[@id='agText']");

            IWebElement firstColumn = null, secondColumn = null;

            foreach (IWebElement column in columnHeaderColl)
                if (column.Text.ToLower().Contains(columnOne.ToLower()))
                {
                    firstColumn = column;
                    break;
                }
            Assert.AreNotEqual(null, firstColumn, "'" + columnOne + "' column not found.");

            foreach (IWebElement column in columnHeaderColl)
                if (column.Text.ToLower().Contains(columnTwo.ToLower()))
                {
                    secondColumn = column;
                    break;
                }
            Assert.AreNotEqual(null, secondColumn, "'" + columnTwo + "' column not found.");

            Actions action = new Actions(driver);
            action.DragAndDrop(firstColumn, secondColumn).Perform();
            Thread.Sleep(1000);

            Results.WriteStatus(test, "Pass", "Verified, Drag And Drop Functionality On Saved Results Columns");
            return new MyReports(driver, test);
        }

        /// <summary>
        /// Verify and Run Selected Type saved Results
        /// </summary>
        /// <param name="savedType">Saved Type</param>
        /// <param name="option">Option Name</param>
        /// <returns></returns>
        public string verifyAndRunSelectedTypeSavedResults(string savedType = "Product Detail", string option = "")
        {
            string searchName = "";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid=0]//div[@class]"), "Saved Results Options icons not present");
            IList<IWebElement> savedResultsOptionsIcon = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid=0]//div[@class]");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='RecordsetType']"), "Saved Results Type name not present");
            IList<IWebElement> savedResultsTypeList = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='RecordsetType']");

            int rowNo = 0;

            for (int i = 0; i < savedResultsTypeList.Count(); i++)
            {
                if (savedResultsTypeList[i].Text == savedType)
                {
                    rowNo = i; break;
                }
            }

            driver._clickByJavaScriptExecutor("//div[@class='tab-pane ng-scope active']//div[@row='" + rowNo + "']/div[@colid=0]//div[@class]");
            IList<IWebElement> optionsDDLList = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + rowNo + "]/div[@colid='0']//li");
            for (int j = 0; j < optionsDDLList.Count(); j++)
                if (optionsDDLList[j].Text == option)
                {
                    driver._clickByJavaScriptExecutor("//div[@class='tab-pane ng-scope active']//div[@row=" + rowNo + "]/div[@colid='0']//li[" + (j + 1) + "]");
                    break;
                }

            Results.WriteStatus(test, "Pass", "Verified, '" + option + "' - '" + savedType + "' Type from Saved Results list.");
            return searchName;
        }

        /// <summary>
        /// Verify Saved Results list - Created By - Search name from list
        /// </summary>
        /// <param name="createdByMe">Verify Created by search</param>
        /// <param name="option">Option name</param>
        /// <returns></returns>
        public string verifySavedResultsList_CreatedBy_SearchName_FromList(bool createdByMe = false, string option = "")
        {
            string searchName = "";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid=0]//div[@class]"), "Saved Search Options icons not present");
            IList<IWebElement> savedSearchOptionsIconColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid=0]//div[@class]");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SavedAs']"), "Saved Name List not present");
            IList<IWebElement> searchNameList = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SavedAs']");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SharedType']"), "Saved Search Options icons not present");
            IList<IWebElement> createByList = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SharedType']");

            for (int i = 0; i < createByList.Count(); i++)
            {
                string[] optionsNameList = new string[] { "Run", "Manage Label" };
                bool avail = false;
                if (createdByMe)
                    if (createByList[i].Text == "Me")
                    {
                        Array.Resize(ref optionsNameList, 3);
                        optionsNameList[1] = "Delete";
                        optionsNameList[2] = "Manage Label";
                    }

                searchName = searchNameList[i].Text;
                driver._clickByJavaScriptExecutor("//div[@class='tab-pane ng-scope active']//div[@row='" + i + "']/div[@colid=0]//div[@class]");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + i + "]/div[@colid='0']//li"), "Saved Search option DDL not present.");
                IList<IWebElement> optionsDDLColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + i + "]/div[@colid='0']//li");

                for (int j = 0; j < optionsDDLColl.Count(); j++)
                {
                    Assert.AreEqual(optionsDDLColl[j].Text, optionsNameList[j], "'" + optionsNameList[j] + "' Option name not present.");
                    if (option == optionsDDLColl[j].Text)
                    {
                        driver._clickByJavaScriptExecutor("//div[@class='tab-pane ng-scope active']//div[@row=" + i + "]/div[@colid='0']//li[" + (j + 1) + "]");
                        avail = true;
                        break;
                    }
                }

                if (avail)
                    break;
            }

            Results.WriteStatus(test, "Pass", "Verified, Created By Options Section.");
            return searchName;
        }

        /// <summary>
        /// Verify Search name Present or Not from Saved Results list
        /// </summary>
        /// <param name="searchName">Search Name</param>
        /// <param name="present">Availble or not</param>
        /// <returns></returns>
        public MyReports verifySavedResultNamePresentOrNotOnList(string tabName, string searchName, bool present = false)
        {
            string colId = "SavedAs";
            if (tabName == "Subscriptions")
                colId = "SubscriptionName";
            if (tabName == "Exported Reports" || tabName.Contains("Best Practice"))
                colId = "ReportName";

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='" + colId + "']"), "Saved Result List not present");
            IList<IWebElement> searchNameList = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='" + colId + "']");

            bool checkingAvail = false;
            for (int i = 0; i < searchNameList.Count(); i++)
                if (searchNameList[i].Text == searchName)
                {
                    checkingAvail = true;
                    break;
                }

            if (present)
                Assert.AreEqual(true, checkingAvail, "'" + searchName + "' " + tabName + " Name not Present on list.");
            else
                Assert.AreEqual(false, checkingAvail, "'" + searchName + "' " + tabName + " not Deleted from list.");

            Results.WriteStatus(test, "Pass", "Verified, " + tabName + " Present or Not from Saved Results List.");
            return new MyReports(driver, test);
        }

        #endregion

        #region Subscription

        /// <summary>
        /// Click Create Subscription Button
        /// </summary>
        /// <returns></returns>
        public MyReports clickCreateSubscriptionButton()
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='tblCreateSub']"), "'Create Subscription' Button not present.");
            driver._clickByJavaScriptExecutor("//input[@id='tblCreateSub']"); Thread.Sleep(1000);
            driver._waitForElementToBePopulated("xpath", "//div[@id='formChkLst']//div[@class='modal-content']");
            Results.WriteStatus(test, "Pass", "Clicked, Create Subscription button.");
            return new MyReports(driver, test);
        }

        /// <summary>
        /// Verify Subscription Section
        /// </summary>
        /// <param name="noRecords">Check No Records message</param>
        /// <returns></returns>
        public MyReports verifySubscriptionSectionInDetail(bool noRecords = false)
        {
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='text-box-input']"), "Search Area not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='text-box-input' and @placeholder='Search']"), "'Search' watermark not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='tblCreateSub']"), "'Create Subscription' Button not present.");

            if (noRecords)
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[contains(@ng-style,'myRptCntrl')]//*[contains(text(),'No records found.')]"), "'No records found.' message not present.");
            else
            {
                IList<IWebElement> columnHeaderColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@id='borderLayout_eRootPanel']//div[@class='ag-header-container']//div[@class='ag-header-row']/div");
                string[] columnHeaderNamesList = new string[] { "Subscription Name", "Type", "Expiration Date", "Last Run", "Created" };

                foreach (string columnHeaderName in columnHeaderNamesList)
                {
                    bool avail = false;
                    foreach (IWebElement columnHeader in columnHeaderColl)
                        if (columnHeader.Text.ToLower().Equals(columnHeaderName.ToLower()))
                        {
                            avail = true;
                            break;
                        }
                    Assert.IsTrue(avail, "'" + columnHeaderName + "' column header name not found.");
                }
            }

            Results.WriteStatus(test, "Pass", "Verified, Subscription Section in Detail.");
            return new MyReports(driver, test);
        }

        /// <summary>
        /// Verify or Select Options of Subscription Name
        /// </summary>
        /// <returns></returns>
        public String verify_OR_Select_OptionsOfSubscriptionName(bool select = false, string subOptionName = "View / Edit")
        {
            string searchName = "";
            if (subOptionName == "Random")
                subOptionName = "View / Edit";
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid=0]//div[@class]"), "Subscription Name Option icons not present");
            IList<IWebElement> subscriptionOptionsIcons = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid=0]//div[@class]");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SubscriptionName']"), "Subscription Name not present");
            IList<IWebElement> subscriptionNamesList = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row]/div[@colid='SubscriptionName']");

            string[] optionsNameList = new string[] { "View / Edit", "Delete", "Send Me Now", "Send To All" };
            bool selectAvail = false;
            for (int l = 0; l < subscriptionNamesList.Count; l++)
            {
                driver._scrollintoViewElement("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + l + "]/div[@colid=0]");
                searchName = subscriptionNamesList[l].Text;
                driver._clickByJavaScriptExecutor("//div[@class='tab-pane ng-scope active']//div[@row='" + l + "']/div[@colid=0]//div[@class]");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + l + "]/div[@colid='0']//li"), "Saved Search option DDL not present.");
                IList<IWebElement> optionsDDLColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@row=" + l + "]/div[@colid='0']//li");

                foreach (string optionName in optionsNameList)
                {
                    bool avail = false;
                    foreach (IWebElement optionDDLEle in optionsDDLColl)
                    {
                        if (select)
                        {
                            if (optionDDLEle.Text.ToLower().Contains(subOptionName.ToLower()))
                            {
                                optionDDLEle.Click(); Thread.Sleep(1000);
                                avail = true;
                                selectAvail = true; break;
                            }
                        }
                        else if (optionDDLEle.Text.ToLower().Contains(optionName.ToLower()))
                        {
                            avail = true;
                            break;
                        }
                    }
                    Assert.IsTrue(avail, "'" + optionName + "' not found in Left Navigation Menu.");
                    if (selectAvail)
                        break;
                }
                if (selectAvail)
                    break;
            }

            Results.WriteStatus(test, "Pass", "Verified, Options Of Subscriptions Name List.");
            return searchName;
        }

        /// <summary>
        /// Verify View / Edit Subscription screen
        /// </summary>
        /// <param name="subName">Subscription Name</param>
        /// <returns></returns>
        public MyReports VerifyViewEditSubscriptionScreen(string subName)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='tabs-1']"), "Report Options tab not present.");
            Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@id='tabs-1']", "class").Contains("active"), "Report Options tab not active.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@id,'selctedQueryName')]"), "Selected Search Name is not present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//span[@id='ctl00_cpHolder_SelectQueryDiv']"), "Search Icon not present beside Search Name.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//input[contains(@id,'ctl00_cpHolder_txtSubName')]"), "Subscription Name not present.");
            Assert.AreEqual(subName, driver._getValue("xpath", "//input[contains(@id,'ctl00_cpHolder_txtSubName')]"), "Subscription Name : " + subName + " is not correct.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//input[contains(@id,'txtSrchReportName')]"), "Search Report Name field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//tr[contains(@id,'TemplateList')]"), "Search Report Name list not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//a[@id='ctl00_cpHolder_lnCreateReportGroup']"), "'Create Report Group' Link not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//a[@id='ctl00_cpHolder_lnDeleteTemplate']"), "'Delete' link not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='ctl00_cpHolder_btnSave2']"), "'Save' Button icon not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='btnSaveAndClose2']"), "'Save & Close' Button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='btnCancel2']"), "'Cancel' Button not present.");

            Results.WriteStatus(test, "Pass", "Verified View / Edit Subscription screen.");
            return new MyReports(driver, test);
        }

        /// <summary>
        /// Select Saved Search Query from Subscription popup
        /// </summary>
        /// <param name="searchName">Search Name to verify</param>
        /// <param name="checkAva">Check availability of Query</param>
        /// <param name="select">Select (true) OR Verify (false)</param>
        /// <returns></returns>
        public String selectSavedSearchQueryFromSubscriptionPopup(string searchName = "Random", bool checkAva = true, bool select = false)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//input[@id='txtSearch']"), "'Search' field not present.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//tr[@class='ng-scope']"), "'Saved Searches' List not present.");
            IList<IWebElement> searchCollection = driver._findElements("xpath", "//div[@class='modal-body-filters']//tr[@class='ng-scope']");
            string queryName = "";

            if (searchName == "Random")
            {
                Random rand = new Random();
                int x = rand.Next(0, searchCollection.Count);
                queryName = searchCollection[x].Text;
                driver._clickByJavaScriptExecutor("//div[@class='modal-body-filters']//tr[@class='ng-scope'][" + (x + 1) + "]/td[1]"); Thread.Sleep(1000);

                Assert.AreEqual("C", driver._getAttributeValue("xpath", "//div[@class='modal-body-filters']//tr[" + (x + 1) + "]/td[2]", "class"), "'" + queryName + "' Query not selected.");
                Results.WriteStatus(test, "Pass", "Selected, '" + queryName + "' Query from Create Subscription Popup.");
            }
            else
            {
                driver._type("xpath", "//div[@class='modal-body-filters']//input[@id='txtSearch']", searchName);
                Thread.Sleep(1000);

                bool avail = false;
                foreach (IWebElement search in searchCollection)
                    if (search.Text.ToLower().Contains(searchName.ToLower()))
                    {
                        avail = true;
                        if (select)
                        {
                            IList<IWebElement> checkBox = search._findElementsWithinElement("xpath", ".//td[@v]");
                            Assert.AreEqual(1, checkBox.Count, "Checkbox not present for saved search '" + search.Text + "'.");
                            checkBox[0].Click();
                            Thread.Sleep(1000);
                        }
                        break;
                    }

                if (checkAva)
                {
                    Assert.IsTrue(avail, "'" + searchName + "' saved search not present.");
                    Results.WriteStatus(test, "Pass", "'" + searchName + "' saved search is present.");
                }
                else
                {
                    Assert.IsFalse(avail, "'" + searchName + "' saved search is present.");
                    Results.WriteStatus(test, "Pass", "'" + searchName + "' saved search not present.");
                }

                Results.WriteStatus(test, "Pass", "Searched, Saved Search from Create Subscription Popup");
            }

            return queryName;
        }

        /// <summary>
        /// Click Column Header for Sort Grid and Verify Column Header Icon
        /// </summary>
        /// <param name="columnName">Column Name for Click and Verify</param>
        /// <param name="ascendingOrder">Verify Column in Ascending Order Else Descending Order</param>
        /// <returns></returns>
        public MyReports clickColumnHeaderForSortGridAndVerifyColumnHeaderIcon(string columnName, bool ascendingOrder)
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[@id='borderLayout_eRootPanel']//div[@class='ag-header-container']//div[@class='ag-header-row']/div"), "Columns not Present.");
            IList<IWebElement> columnCollections = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@id='borderLayout_eRootPanel']//div[@class='ag-header-container']//div[@class='ag-header-row']/div");

            string commonPath = "//div[@class='tab-pane ng-scope active']//div[@id='borderLayout_eRootPanel']//div[@class='ag-header-container']//div[@class='ag-header-row']/div";

            for (int i = 0; i < columnCollections.Count; i++)
            {
                if (columnCollections[i].Text.Trim().Replace("\r\n", "").Contains(columnName))
                {
                    Assert.AreEqual(true, driver._isElementPresent("xpath", "" + commonPath + "[" + (i + 1) + "]/div[@id='agHeaderCellLabel']/span[@id='agText']"), "Column Text area not Present for '" + columnName + "' Column.");

                    if (ascendingOrder)
                    {
                        if (driver._isElementPresent("xpath", "" + commonPath + "[" + (i + 1) + "]/div[@id='agHeaderCellLabel']/span[@id='agSortAsc']") == false)
                        {
                            driver._clickByJavaScriptExecutor("" + commonPath + "[" + (i + 1) + "]/div[@id='agHeaderCellLabel']/span[@id='agText']");
                            Thread.Sleep(1000);
                            driver._waitForElementToBeHidden("xpath", "//div[@class='ProcessLoader']");
                            Results.WriteStatus(test, "Pass", "Clicked, '" + columnName + "' Column to Sort Grid in Ascending Order.");
                        }

                        Assert.AreEqual(true, driver._isElementPresent("xpath", "" + commonPath + "[" + (i + 1) + "]/div[@id='agHeaderCellLabel']/span[@id='agSortAsc' and @class='ag-header-icon ag-sort-ascending-icon']"), "Up Arrow for '" + columnName + "' Column not Present when Column Sort in Ascending Order.");
                        Results.WriteStatus(test, "Pass", "Verified, Up Arrow Icon for '" + columnName + "' Column Header.");
                    }
                    else
                    {
                        if (driver._isElementPresent("xpath", "" + commonPath + "[" + (i + 1) + "]/div[@id='agHeaderCellLabel']/span[@id='agSortDesc']") == false)
                        {
                            driver._clickByJavaScriptExecutor("" + commonPath + "[" + (i + 1) + "]/div[@id='agHeaderCellLabel']/span[@id='agText']");
                            Thread.Sleep(1000);
                            driver._waitForElementToBeHidden("xpath", "//div[@class='ProcessLoader']");
                            Results.WriteStatus(test, "Pass", "Clicked, '" + columnName + "' Column to Sort Grid in Descending Order.");
                        }

                        Assert.AreEqual(true, driver._isElementPresent("xpath", "" + commonPath + "[" + (i + 1) + "]/div[@id='agHeaderCellLabel']/span[@id='agSortDesc' and @class='ag-header-icon ag-sort-descending-icon']"), "Down Arrow for '" + columnName + "' Column not Present when Column Sort in Descending Order.");
                        Results.WriteStatus(test, "Pass", "Verified, Down Arrow Icon for '" + columnName + "' Column Header.");
                    }
                    break;
                }
            }
            return new MyReports(driver, test);
        }

        #endregion

        #region Scorecard Chart

        /// <summary>
        /// Verify Scorecard Chart Section
        /// </summary>
        /// <param name="noRecords">when No Records avail</param>
        /// <returns></returns>
        public MyReports verifyScorecardChartSectionInDetail(bool noRecords = false)
        {
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='text-box-input']"), "Search Area not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='text-box-input' and @placeholder='Search']"), "'Search' watermark not Present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//input[@id='tblCreateSub']"), "'Create Subscription' Button not present.");

            if (noRecords)
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='tab-pane ng-scope active']//div[contains(@ng-style,'myRptCntrl')]//*[contains(text(),'No records found.')]"), "'No records found.' message not present.");
            else
            {
                IList<IWebElement> columnHeaderColl = driver._findElements("xpath", "//div[@class='tab-pane ng-scope active']//div[@id='borderLayout_eRootPanel']//div[@class='ag-header-container']//div[@class='ag-header-row']/div");
                string[] columnHeaderNamesList = new string[] { "Subscription Name", "Type", "Expiration Date", "Last Run", "Created" };

                foreach (string columnHeaderName in columnHeaderNamesList)
                {
                    bool avail = false;
                    foreach (IWebElement columnHeader in columnHeaderColl)
                        if (columnHeader.Text.ToLower().Equals(columnHeaderName.ToLower()))
                        {
                            avail = true;
                            break;
                        }
                    Assert.IsTrue(avail, "'" + columnHeaderName + "' column header name not found.");
                }
            }

            Results.WriteStatus(test, "Pass", "Verified, Scorecard Chart Section in Detail.");
            return new MyReports(driver, test);
        }

        #endregion


        #endregion
    }
}