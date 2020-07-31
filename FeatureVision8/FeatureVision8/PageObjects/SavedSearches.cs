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
    public class SavedSearches
    {
        #region Private Variables

        private IWebDriver savedSearches;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public SavedSearches(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.savedSearches = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.savedSearches; }
            set { this.savedSearches = value; }
        }

        ///<summary>
        ///Verify Saved Searches Screen
        ///</summary>
        ///<param name="withList">Whether the table should have saved searches in it</param>
        ///<returns></returns>
        public SavedSearches VerifySavedSearchesScreen(bool withList = true)
        {
            Assert.AreEqual("Saved Searches".ToLower(), homePage.GetActiveScreenNameFromLeftNavigationMenu().ToLower(), "'Saved Searches Screen is not active.'");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='ContentWrapper']"), "Saved Search Screen not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='text-box-input']"), "Search Box not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='tblCreateSub']"), "Create Subscription button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@ag-grid-id]"), "Saved Searches Table not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'ag-header-cell-sortable')]//span[@id='agText']"), "Column Headers not presentin table.");
            IList<IWebElement> columnHeaderColl = driver._findElements("xpath", "//div[contains(@class, 'ag-header-cell-sortable')]//span[@id='agText']");
            string[] columnHeaderNamesList = new string[] { "Promo Search", "Created By", "Label", "Type", "Subs", "Last Run", "Created", "Query Type" };

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

            if (withList)
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-body-container']//div[@row]"), "Rows not present in table.");
            else
            {
                if (driver._isElementPresent("xpath", "//div[@class='ag-body-container']//div[@row]"))
                    Results.WriteStatus(test, "Info", "Saved Searches list present.'No records found' message not verify");
                else
                {
                    Assert.IsFalse(driver._isElementPresent("xpath", "//div[@class='ag-body-container']//div[@row]"), "Rows are present in table.");
                    Assert.IsTrue(driver._isElementPresent("xpath", "//*[contains(text(), 'No Record Available')]"), "Rows not present in table.");
                }
            }
            Assert.IsTrue(driver._isElementPresent("xpath", "//button[contains(@ng-click, 'showSearchListDropdown')]"), "Show dropdown not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//button[contains(@ng-click, 'createdSearchListDropdown')]"), "Created By dropdown not present.");

            Results.WriteStatus(test, "Pass", "Verified, Saved Searches Screen");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Show Dropdown
        ///</summary>
        ///<param name="option">Option to select from Show Dropdown</param>
        ///<returns></returns>
        public SavedSearches VerifyShowDropdown(string option)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[text()='Show:']"), "Show dropdown label not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//button[contains(@ng-click, 'showSearchListDropdown')]"), "Show dropdown not present.");

            driver._click("xpath", "//button[contains(@ng-click, 'showSearchListDropdown')]");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@ng-class, 'showSearchListDropdown')]//li"), "Show DDL not present.");
            IList<IWebElement> showDDLColl = driver._findElements("xpath", "//div[contains(@ng-class, 'showSearchListDropdown')]//li");

            foreach (IWebElement showEle in showDDLColl)
                if (showEle.Text.ToLower().Contains(option.ToLower()))
                {
                    showEle.Click();
                    break;
                }

            Assert.AreEqual(option.ToLower(), driver._getText("xpath", "//button[contains(@ng-click, 'showSearchListDropdown')]").ToLower(), "'" + option + "' not selected");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@row]/div[@colid='SavedQueryType']"), "Cells not present in Query type column.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//div[@row]/div[@colid='SavedQueryType']");

            if (option.ToLower().Equals("shared"))
            {
                int i = 0;
                foreach (IWebElement cell in cellCollection)
                {
                    driver._scrollintoViewElement("xpath", "//div[@row=" + (i++) + "]/div[@colid='SavedQueryType']");
                    Assert.AreEqual("public", cell.Text.ToLower(), "'Shared' option was not applied successfully.");
                }
            }
            else
            {
                foreach (IWebElement cell in cellCollection)
                    Assert.IsTrue(cell.Text.ToLower().Equals("public") || cell.Text.ToLower().Equals("private"), "'Shared' option was not applied successfully.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Show Dropdown");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Created By Dropdown
        ///</summary>
        ///<param name="option">Option to select from Show Dropdown</param>
        ///<returns></returns>
        public SavedSearches VerifyCreatedByDropdown(string option)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[text()='Created By:']"), "Show dropdown label not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//button[contains(@ng-click, 'createdSearchListDropdown')]"), "Created By dropdown not present.");

            driver._click("xpath", "//button[contains(@ng-click, 'createdSearchListDropdown')]");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@ng-class, 'createdSearchListDropdown')]//li"), "Show DDL not present.");
            IList<IWebElement> showDDLColl = driver._findElements("xpath", "//div[contains(@ng-class, 'createdSearchListDropdown')]//li");

            foreach (IWebElement showEle in showDDLColl)
                if (showEle.Text.ToLower().Contains(option.ToLower()))
                {
                    showEle.Click();
                    break;
                }

            Assert.AreEqual(option.ToLower(), driver._getText("xpath", "//button[contains(@ng-click, 'createdSearchListDropdown')]").ToLower(), "'" + option + "' not selected");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@row]/div[@colid='SharedType']"), "Cells not present in Query type column.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//div[@row]/div[@colid='SharedType']");

            if (option.ToLower().Equals("me"))
            {
                int i = 0;
                foreach (IWebElement cell in cellCollection)
                {
                    driver._scrollintoViewElement("xpath", "//div[@row=" + (i++) + "]/div[@colid='SharedType']");
                    Assert.AreEqual("me", cell.Text.ToLower(), "'Shared' option was not applied successfully.");
                }
            }
            else if (option.ToLower().Equals("client"))
            {
                int i = 0;
                foreach (IWebElement cell in cellCollection)
                {
                    driver._scrollintoViewElement("xpath", "//div[@row=" + (i++) + "]/div[@colid='SharedType']");
                    Assert.AreEqual("client", cell.Text.ToLower(), "'Shared' option was not applied successfully.");
                }
            }
            else if (option.ToLower().Equals("numerator"))
            {
                int i = 0;
                foreach (IWebElement cell in cellCollection)
                {
                    driver._scrollintoViewElement("xpath", "//div[@row=" + (i++) + "]/div[@colid='SharedType']");
                    Assert.AreEqual("numerator", cell.Text.ToLower(), "'Shared' option was not applied successfully.");
                }
            }

            Results.WriteStatus(test, "Pass", "Verified, Created By Dropdown");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Search Box
        ///</summary>
        ///<param name="searchName">To search from Search Box</param>
        ///<returns></returns>
        public SavedSearches VerifySearchBox(string searchName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='text-box-input']"), "Search Box not present.");
            Assert.AreEqual("Promo Search", driver._getAttributeValue("xpath", "//input[@id='text-box-input']", "placeholder"), "'Search Box' placeholder text does not match.");

            driver._type("xpath", "//input[@id='text-box-input']", searchName);
            Thread.Sleep(1000);

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@row]/div[@colid='QueryName']/a"), "Cells not present in Query type column.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//div[@row]/div[@colid='QueryName']/a");

            for (int i = 0; i < cellCollection.Count; i++)
            {
                driver._scrollintoViewElement("xpath", "//div[@row=" + i + "]/div[@colid='QueryName']");
                Assert.IsTrue(cellCollection[i].Text.ToLower().Contains(searchName.ToLower()), "'" + searchName + "' Saved Search was not applied successfully.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Created By Dropdown");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Saved Search Options
        ///</summary>
        ///<param name="option">Option to click</param>
        ///<returns></returns>
        public string VerifySavedSearchOptions(string option = "", bool createdByMe = false)
        {
            string searchName = "";
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@row]/div[@colid=0]//div[@class]"), "Saved Search Options icons not present");
            IList<IWebElement> savedSearchOptionsIconColl = driver._findElements("xpath", "//div[@row]/div[@colid=0]//div[@class]");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@row]/div[@colid='QueryName']/a"), "Saved Search Options icons not present");
            IList<IWebElement> savedSearchNameColl = driver._findElements("xpath", "//div[@row]/div[@colid='QueryName']/a");

            Random rand = new Random();
            int x = rand.Next(0, savedSearchOptionsIconColl.Count);

            driver._scrollintoViewElement("xpath", "//div[@row=" + x + "]/div[@colid=0]");
            searchName = savedSearchNameColl[x].Text;

            savedSearchOptionsIconColl[x].Click();
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@row=" + x + "]/div[@colid='0']//li"), "Saved Search option DDL not present.");
            IList<IWebElement> optionsDDLColl = driver._findElements("xpath", "//div[@row=" + x + "]/div[@colid='0']//li");
            string[] optionsNameList = new string[] { "Run", "View Criteria", "Manage Label" };
            if (createdByMe)
            {
                Array.Resize(ref optionsNameList, 4);
                optionsNameList[2] = "Delete";
                optionsNameList[3] = "Manage Label";
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
        ///Verify View Criteria Window
        ///</summary>
        ///<returns></returns>
        public SavedSearches VerifyViewCriteriaWindow()
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//table[@id='popupDiv1']"), "View Criteria Window not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvTitle1']"), "View Criteria Window Header not present.");
            Assert.AreEqual("View Criteria", driver._getText("xpath", "//div[@id='dvTitle1']"), "View Criteria Window Header text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//table//td[@class]"), "View Criteria Details Paramenters not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table//td[@width and text()]"), "View Criteria Window Details Parameter Values not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//table//div[@class='close']"), "View Criteria Window Cross button not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//table//input[@value='Cancel']"), "View Criteria Window Close button not present.");

            Results.WriteStatus(test, "Pass", "Verified, View Criteria Window");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Manage Label Options Popup
        ///</summary>
        ///<param name="searchName">Saved Search Name</param>
        ///<param name="popupVisible">Whether popup should be visible</param>
        ///<returns></returns>
        public SavedSearches VerifyManageLabelOptionsPopup(string searchName, bool popupVisible = true)
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Manage Label Options Popup not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//h4"), "Manage Label Options Popup header not present.");
                Assert.AreEqual("Labels for " + searchName, driver._getText("xpath", "//div[@class='modal-content']//h4"), "Manage Label Options Popup header text does not match.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//input[@type='text']"), "Add New Label Name Field not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[@title='Add New Label']"), "Add New Label Name Button not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[@id='availLabels']"), "Available Label(s) Section not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[@id='selLabels']"), "Selected Label(s) Section not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'label-navigate')]//div[contains(@ng-click,'ToReport')]"), "'Assign Label To Report' Arrow not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'label-navigate')]//div[contains(@ng-click,'FromReport')]"), "'Remove Label From Report' Arrow not present.");

                if (driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//table[@id='tblNotAssign']//tr") == false)
                {
                    string labelName = "Test" + driver._randomString(3, true);
                    driver._type("xpath", "//div[@class='modal-body-filters']//input[@type='text']", labelName); Thread.Sleep(500);
                    driver._click("xpath", "//div[@class='modal-body-filters']//div[@title='Add New Label']"); Thread.Sleep(500);
                    Results.WriteStatus(test, "Pass", "Added New Label name.");
                }

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//table[@id='tblNotAssign']//tr"), "Available Label Names List not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[@id='selLabels']"), "Selected Label Names List not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-footer']//button[text()='Ok']"), "Ok Button not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-footer']//button[text()='Cancel']"), "Cancel button not present.");
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "Manage Label Options Popup is present.");
                Results.WriteStatus(test, "Pass", "Verified, Manage Label Options Popup is closed.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Manage Label Options Popup");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Change Label From Manage Label Options
        ///</summary>
        ///<returns></returns>
        public string ChangeLabelFromManageLabelOptions()
        {
            string newLabel = "";
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//table[@id='tblNotAssign']//tr"), "Available Label Names List not present.");
            IList<IWebElement> availLabelsColl = driver._findElements("xpath", "//div[@class='modal-body-filters']//table[@id='tblNotAssign']//tr");

            Random rand = new Random();
            int x = rand.Next(0, availLabelsColl.Count);

            IList<IWebElement> newLabelText = availLabelsColl[x]._findElementsWithinElement("xpath", ".//div[contains(@class, 'btnoption')]");
            Assert.AreEqual(1, newLabelText.Count, "Label Name text not present.");
            newLabel = newLabelText[0].Text;
            newLabelText[0].Click();
            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='tblNotAssign']//div[contains(@class, 'selected') and text()='" + newLabel + "']"), "Label not selected.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'long-arrow-right')]"), "Select Arrow not present.");
            driver._click("xpath", "//div[contains(@class, 'long-arrow-right')]");

            Assert.IsTrue(driver._isElementPresent("xpath", "//table[@id='tblAssign']//div[contains(@class, 'btnoption')]"), "Available Label Names List not present.");
            IList<IWebElement> selectedLabelsColl = driver._findElements("xpath", "//table[@id='tblAssign']//div[contains(@class, 'btnoption')]");

            bool avail = false;
            foreach (IWebElement selectedLabel in selectedLabelsColl)
                if (selectedLabel.Text.ToLower().Equals(newLabel.ToLower()))
                {
                    avail = true;
                    break;
                }
            Assert.IsTrue(avail, "'" + newLabel + "' Label not selected.");

            driver._click("xpath", "//div[@class='modal-footer']//button[text()='Ok']");

            Results.WriteStatus(test, "Pass", "Changed, Label to '" + newLabel + "' From Manage Label Options.");
            return newLabel;
        }

        ///<summary>
        ///Verify Label of Saved Searches
        ///</summary>
        ///<param name="labelName">Label Name to search</param>
        ///<returns></returns>
        public SavedSearches VerifyLabelOfSavedSearches(string labelName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='text-box-input']"), "Search Box not present.");
            Assert.AreEqual("Promo Search", driver._getAttributeValue("xpath", "//input[@id='text-box-input']", "placeholder"), "'Search Box' placeholder text does not match.");

            driver._type("xpath", "//input[@id='text-box-input']", labelName);
            Thread.Sleep(1000);

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@row]/div[@colid='LabelText']"), "Cells not present in Query type column.");
            IList<IWebElement> cellCollection = driver._findElements("xpath", "//div[@row]/div[@colid='LabelText']");

            for (int i = 0; i < cellCollection.Count; i++)
            {
                driver._scrollintoViewElement("xpath", "//div[@row=" + i + "]/div[@colid='LabelText']");
                Assert.IsTrue(cellCollection[i].Text.ToLower().Contains(labelName.ToLower()), "'" + labelName + "' Label Name was not applied successfully.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Label of Saved Searches.");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Create Subscription Popup
        ///</summary>
        ///<param name="popupVisible">Whether popup should be visible</param>
        ///<returns></returns>
        public SavedSearches VerifyCreateSubscriptionPopup(bool popupVisible = true)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='tblCreateSub']"), "Create Subscription button not present.");
            driver._click("xpath", "//input[@id='tblCreateSub']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "'Select one or more saved queries to define your subscription' popup not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']//h4"), "'Select one or more saved queries to define your subscription' popup header not present.");
            Assert.AreEqual("Select one or more saved queries to define your subscription", driver._getText("xpath", "//div[@class='modal-content']//h4"), "'Select one or more saved queries to define your subscription' popup header text does not match.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//button[@id='dropdownMenu1']"), "'Shared by' field not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//input[@id='txtSearch']"), "'Search' field not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//tr"), "'Saved Searches' List not present.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'modal-footer')]//button[text()='Continue']"), "'Continue' button not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'modal-footer')]//button[text()='Cancel']"), "'Cancel' button not present.");

            Results.WriteStatus(test, "Pass", "Verified, Create Subscription Popup.");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Search And Select Saved Search In Create Subscription Popup
        ///</summary>
        ///<returns></returns>
        public SavedSearches SearchAndSelectSavedSearchInCreateSubscriptionPopup(string searchName, bool present = true, bool select = false)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//input[@id='txtSearch']"), "'Search' field not present.");
            driver._type("xpath", "//div[@class='modal-body-filters']//input[@id='txtSearch']", searchName);
            Thread.Sleep(1000);

            if (driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//tr[@class='ng-scope']"))
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//tr[@class='ng-scope']"), "'Saved Searches' List not present.");
                IList<IWebElement> searchCollection = driver._findElements("xpath", "//div[@class='modal-body-filters']//tr[@class='ng-scope']");

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

                if (present)
                {
                    Assert.IsTrue(avail, "'" + searchName + "' saved search not present.");
                    Results.WriteStatus(test, "Pass", "'" + searchName + "' saved search is present.");
                }
                else
                {
                    Assert.IsFalse(avail, "'" + searchName + "' saved search is present.");
                    Results.WriteStatus(test, "Pass", "'" + searchName + "' saved search not present.");
                }
            }
            else
                if (present == false)
                Results.WriteStatus(test, "Pass", "'" + searchName + "' saved search not present.");

            Results.WriteStatus(test, "Pass", "Searched, Saved Search from Create Subscription Popup");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Drag And Drop Functionality On Saved Search Columns
        ///</summary>
        ///<returns></returns>
        public SavedSearches VerifyDragAndDropFunctionalityOnSavedSearchColumns(string columnOne, string columnTwo)
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

            Results.WriteStatus(test, "Pass", "Verified, Drag And Drop Functionality On Saved Search Columns");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Summary Screen
        ///</summary>
        ///<param name="searchName">Name of Search</param>
        ///<returns></returns>
        public SavedSearches VerifySummaryScreen(string searchName = "")
        {
            Assert.AreEqual("Summary", homePage.GetActiveScreenNameFromLeftNavigationMenu(), "Summary Screen is not Active");



            Results.WriteStatus(test, "Pass", "Verified, Summary Screen");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify And Edit Report Options Tab of Subscription
        ///</summary>
        ///<param name="searchName">Name of Search</param>
        ///<returns></returns>
        public string VerifyReportOptionsAndEditTabOfSubscription(string searchName)
        {
            if (driver._getAttributeValue("xpath", "//div[@data-loader]//div[@class='ProcessLoader']", "style") == "")
                driver._waitForElementToBePopulated("xpath", "//div[@data-loader and contains(@style,'display')]//div[@class='ProcessLoader']");

            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='tabs-1']"), "Report Options tab not present.");
            Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@id='tabs-1']", "class").Contains("active"), "Report Options tab not active.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//span[contains(@id,'selctedQueryName')]"), "Selected Search Name is not present.");
            Assert.AreEqual(searchName, driver._getText("xpath", "//span[contains(@id,'selctedQueryName')]"), "Selected Search Name is not correct.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//input[contains(@id,'txtSubName')]"), "Subscription Name field not present.");
            string subscriptionName = searchName.Substring(0, 6) + "_Sub";
            driver._type("xpath", "//input[contains(@id,'txtSubName')]", subscriptionName);

            Assert.IsTrue(driver._isElementPresent("xpath", "//input[contains(@id,'txtSrchReportName')]"), "Search Report Name field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//tr[contains(@id,'TemplateList')]"), "Search Report Name list not present.");
            string reportName = driver._getText("xpath", "//tr[contains(@id,'TemplateList')][1]/td");
            driver._click("xpath", "//tr[contains(@id,'TemplateList')][1]");

            Results.WriteStatus(test, "Pass", "Verified And Editted, Report Options Tab of Subscription");
            return subscriptionName;
        }

        ///<summary>
        ///Check checkbox of Subscription Screen
        ///</summary>
        ///<param name="checkbox">Name of chdckbox to be checked</param>
        ///<returns></returns>
        public SavedSearches CheckCheckboxOfSubscriptionScreen(string checkbox)
        {
            if (checkbox.ToLower().Contains("calendar"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id,'chkIncludeCalendar')]"), "'" + checkbox + "' checkbox not present.");
                driver._click("xpath", "//div[contains(@id,'chkIncludeCalendar')]");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[contains(@id,'chkIncludeCalendar')]", "class").Contains("checked"), "'" + checkbox + "' checkbox not checked.");
            }
            else if (checkbox.ToLower().Contains("product/ad block"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id,'_chkIncPI')]"), "'" + checkbox + "' checkbox not present.");
                driver._click("xpath", "//div[contains(@id,'_chkIncPI')]");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[contains(@id,'_chkIncPI')]", "class").Contains("checked"), "'" + checkbox + "' checkbox not checked.");
            }
            else if (checkbox.ToLower().Contains("page image"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id,'_chkIncPgImg')]"), "'" + checkbox + "' checkbox not present.");
                driver._click("xpath", "//div[contains(@id,'_chkIncPgImg')]");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[contains(@id,'_chkIncPgImg')]", "class").Contains("checked"), "'" + checkbox + "' checkbox not checked.");
            }
            else if (checkbox.ToLower().Contains("product information"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id,'_chkProductInfo')]"), "'" + checkbox + "' checkbox not present.");
                driver._click("xpath", "//div[contains(@id,'_chkProductInfo')]");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[contains(@id,'_chkProductInfo')]", "class").Contains("checked"), "'" + checkbox + "' checkbox not checked.");
            }

            Results.WriteStatus(test, "Pass", "Checked, '" + checkbox + "' checkbox of Subscription Screen");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify And Edit Schedule And Format Tab of Subscription
        ///</summary>
        ///<param name="recursive">To select radio button</param>
        ///<returns></returns>
        public SavedSearches VerifyAndEditScheduleAndFormatTabOfSubscription(bool recursive = false)
        {
            if (!driver._getAttributeValue("xpath", "//div[@id='tabs-4']", "class").Contains("active"))
            {
                driver._click("xpath", "//li[contains(@id, 'liTabs4')]/a");
                Thread.Sleep(2000);
            }
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='tabs-4']"), "Schedule And Format tab not present.");
            Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@id='tabs-4']", "class").Contains("active"), "Schedule And Format tab not active.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'rdoOnce')]"), "'Send Once' Radio button not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'rdoRecursive')]"), "'Recursive' Radio button not present");

            if (recursive)
            {
                driver._click("xpath", "//div[contains(@id, 'rdoRecursive')]");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[contains(@id, 'rdoRecursive')]", "class").Contains("actibe"), "Recursive Radio button not selected.");
            }
            else
            {
                driver._click("xpath", "//div[contains(@id, 'rdoOnce')]");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[contains(@id, 'rdoOnce')]", "class").Contains("active"), "Send Once Radio button not selected.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'icontxtStart')]"), "'Start Date' Calendar button not present");
            driver._click("xpath", "//div[contains(@id, 'icontxtStart')]");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@id, 'ui-datepicker-div')]"), "Datepicker calendar not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//td[@class][1]/a"), "Days not present in calendar");
            IList<IWebElement> daysCollection = driver._findElements("xpath", "//td[@class][1]/a");
            daysCollection[0].Click();
            Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[contains(@id, 'ui-datepicker-div')]"), "Datepicker calendar not present.");

            if (recursive)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[@id='txtFreq']"), "Frequency Text field not present.");
                driver._type("xpath", "//input[@id='txtFreq']", "2");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'icontxtEnd')]"), "'End Date' Calendar button not present");
                driver._click("xpath", "//div[contains(@id, 'icontxtEnd')]");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@id, 'ui-datepicker-div')]"), "Datepicker calendar not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//td[@class]/a"), "Days not present in calendar");
                daysCollection = driver._findElements("xpath", "//td[@class]/a");
                daysCollection[daysCollection.Count - 1].Click();
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[contains(@id, 'ui-datepicker-div')]"), "Datepicker calendar not present.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//input[contains(@value, 'Add >>') and not(@id)]"), "Add Button not present.");
            driver._click("xpath", "//input[contains(@value, 'Add >>') and not(@id)]");

            Assert.IsTrue(driver._isElementPresent("xpath", "//tr[contains(@id, '_SchedulesDivr')]"), "Added Schedule not present.");
            driver._click("xpath", "//tr[contains(@id, '_SchedulesDivr')][1]");

            Results.WriteStatus(test, "Pass", "Verified And Editted, Schedule And Format Tab of Subscription");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Click Button On Summary Screen
        ///</summary>
        ///<param name="buttonName">To click button</param>
        ///<returns></returns>
        public SavedSearches ClickButtonOnSummaryScreen(string buttonName)
        {
            if (buttonName.ToLower().Equals("save"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'btnSave2')]"), "'" + buttonName + "' button not present.");
                driver._click("xpath", "//div[contains(@id, 'btnSave2')]");
            }
            else if (buttonName.ToLower().Equals("save & close"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[contains(@id, 'btnSaveAndClose2')]"), "'" + buttonName + "' button not present.");
                driver._click("xpath", "//input[contains(@id, 'btnSaveAndClose2')]");
            }
            else if (buttonName.ToLower().Equals("cancel"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//input[contains(@id, 'btnCancel2')]"), "'" + buttonName + "' button not present.");
                driver._click("xpath", "//input[contains(@id, 'btnCancel2')]");
            }

            Results.WriteStatus(test, "Pass", "Clicked, '" + buttonName + "' Button On Summary Screen");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Subscriptions for Query Popup
        ///</summary>
        ///<returns></returns>
        public SavedSearches VerifySubscriptionsForQueryPopup(string searchName, string[] subsNameList = null)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Subscriptions for Query Popup not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']//h4"), "Subscriptions for Query Popup header not present.");
            Assert.AreEqual("Subscriptions for Query:" + searchName, driver._getText("xpath", "//div[@class='modal-content']//h4"), "Subscriptions for Query Popup header text does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div[@colid]//span[@id='agText']"), "Column Headers not present on Subscription for Query popup.");
            IList<IWebElement> columnColl = driver._findElements("xpath", "//div[@class='modal-content']//div[@colid]//span[@id='agText']");
            string[] columnNameList = new string[] { "Name", "Created", "Last Run", "Expiration Date", "View", "Delete", "Send Me", "Send All" };

            foreach (string columnName in columnNameList)
            {
                bool avail = false;
                foreach (IWebElement column in columnColl)
                    if (column.Text.ToLower().Contains(columnName.ToLower()))
                    {
                        avail = true;
                        break;
                    }
                Assert.IsTrue(avail, "'" + columnName + "' Column not found.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div[@row]/div[@colid='SubscriptionName']"), "'Subscription Name' column cells not present on Subscription for Query popup.");
            IList<IWebElement> subCellColl = driver._findElements("xpath", "//div[@class='modal-content']//div[@row]/div[@colid='SubscriptionName']");

            if (subsNameList != null)
            {
                foreach (string subsName in subsNameList)
                {
                    bool avail = false;
                    foreach (IWebElement cell in subCellColl)
                        if (cell.Text.ToLower().Contains(subsName.ToLower()))
                        {
                            avail = true;
                            break;
                        }
                    Assert.IsTrue(avail, "'" + subsName + "' Column not found.");
                }
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class,'modal-footer')]//button[text()='Cancel']"), "Cancel button not present.");

            Results.WriteStatus(test, "Pass", "Verified, Subscriptions for Query Popup");
            return new SavedSearches(driver, test);
        }

        ///<summary>
        ///Verify Subs Column on Saved Searches Screen
        ///</summary>
        ///<returns></returns>
        public SavedSearches VerifySubsColumnOnSavedSearchesScreen(string searchName, string subsNum = "1", bool clickSubs = false)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@row]/div[@colid='QueryName']/a"), "Cells not present in Query Name column.");
            IList<IWebElement> nameCollection = driver._findElements("xpath", "//div[@row]/div[@colid='QueryName']/a");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@row]/div[@colid='SubscriptionCount']/a"), "Cells not present in Subs column.");

            for (int i = 0; i < nameCollection.Count; i++)
            {
                driver._scrollintoViewElement("xpath", "//div[@row=" + i + "]/div[@colid='QueryName']");
                if (nameCollection[i].Text.ToLower().Contains(searchName.ToLower()))
                {
                    Assert.AreEqual(subsNum, driver._getText("xpath", "//div[@row=" + i + "]/div[@colid='SubscriptionCount']/a"), "Subs count of Saved Search '" + searchName + "' does not match");
                    if (clickSubs)
                        driver._click("xpath", "//div[@row=" + i + "]/div[@colid='SubscriptionCount']/a");
                    Results.WriteStatus(test, "Pass", "Verified, Subs Count as '" + subsNum + "' for Saved Searches '" + searchName + "'");
                    break;
                }
            }

            Results.WriteStatus(test, "Pass", "Verify Subs Column on Saved Searches Screen");
            return new SavedSearches(driver, test);
        }

        #region Vishal New Methods

        /// <summary>
        /// Get Saved Searches name from list
        /// </summary>
        /// <returns></returns>
        public String getSavedSearchNameFromList()
        {
            string savedSearchName = "";
            if (driver._isElementPresent("xpath", "//div[@class='ag-body-container']//div[@row]"))
                savedSearchName = driver._getText("xpath", "//div[@class='ag-body-container']//div[@row]//div[@colid='QueryName']");
            else
                Assert.IsFalse(driver._isElementPresent("xpath", "//div[@class='ag-body-container']//div[@row]//div[@colid='QueryName']"), "Saved Searches list not present.");

            Results.WriteStatus(test, "Pass", "Verified Saved Searches list and get Saved Search Name : " + savedSearchName + ".");
            return savedSearchName;
        }

        #endregion

        #endregion
    }
}
