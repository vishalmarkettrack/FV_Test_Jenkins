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
    public class Home
    {
        #region Private Variables

        private IWebDriver home;
        private ExtentTest test;

        #endregion

        #region Public Methods

        public Home(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.home = driver;
            test = testReturn;
        }

        public IWebDriver driver
        {
            get { return this.home; }
            set { this.home = value; }
        }

        ///<summary>
        ///Verify Home Page
        ///</summary>
        ///<returns></returns>
        public Home VerifyHomePage()
        {
            driver._waitForElementToBeHidden("xpath", "//div[@class='ProcessLoader']", 65);
            driver._waitForElementToBeHidden("xpath", "//div[contains(@class,'PopupLoader')]", 65);

            if (driver._isElementPresent("xpath", "//div[@class='PopupDiv ui-draggable']"))
            {
                Assert.AreEqual(true, driver._isElementPresent("id", "chkDoNotShowThisMessageAgain"), "'Do not show this announcement again.' Checkbox not present.");
                driver._click("id", "chkDoNotShowThisMessageAgain");
                Thread.Sleep(500);
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@id='chkDoNotShowThisMessageAgain' and contains(@class,'active')]"), "'Do not show this announcement again.' Checkbox not checked.");

                Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='PopupDiv ui-draggable']//.//input[@type='button' and @value='Close']"), "'Close' Button not present.");
                driver._clickByJavaScriptExecutor("//div[@class='PopupDiv ui-draggable']//.//input[@type='button' and @value='Close']");
                Thread.Sleep(1000);
            }

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//img[@src='/Images/isc-NumeratorLogo.png']"), "Numerator Logo not found on Home Page.");
            driver._waitForElementToBeHidden("id", "ctl00_ctlLoader_divProcessing", 45);
            driver._waitForElementToBeHidden("xpath", "//div[@class='LoaderStyle LoaderContainer']", 45);

            Results.WriteStatus(test, "Pass", "Verified, Home Page");
            return new Home(driver, test);
        }

        /// <summary>
        /// To Verify Home Screen Detail
        /// </summary>
        /// <returns></returns>
        public Home VerifyHomeScreenInDetail(string searchName = "", bool DetailLevel = true, bool AdLevel = false)
        {
            driver._waitForElementToBeHidden("xpath", "//div[@class='ProcessLoader'], 45");
            Assert.IsTrue(driver._waitForElement("xpath", "//img[@src='/Images/isc-NumeratorLogo.png']", 15), "Numerator Logo not found on Home Page.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='report-content']/h1"), "Madlib Search header text not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[contains(@class, 'filler')]"), "Madlib Prompt Summary text not present.");

            if (searchName != "")
                Assert.IsTrue(driver._getText("xpath", "//div[@class='report-content']/h1").ToLower().Contains(searchName.ToLower()), "Search Name '" + searchName + "' does not match.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='1']"), "Madlib Search Parameter 'Any Product' not present.");
            if (driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='5']"))
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='5']"), "Madlib Search Parameter 'Any Specific Account/Market' not present.");
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='2']"), "Madlib Search Parameter 'Any Retailer' not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='3']"), "Madlib Search Parameter 'Any Market' not present.");
            }
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid='4']"), "Madlib Search Parameter 'Any Date' not present.");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@ui-view='Header']//li"), "Tabs not present.");

            if (DetailLevel)
            {
                string[] tabNameList = new string[] { "Promoted Products", "Ad Blocks", "Pages", "Ads" };
                IList<IWebElement> tabCollection = driver._findElements("xpath", "//div[@ui-view='Header']//li");

                if (AdLevel)
                {
                    tabNameList[0] = "Pages";
                    tabNameList[1] = "Ads";
                    Array.Resize(ref tabNameList, 2);
                }

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
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Multi-select']/i"), "'Multiselect' option not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Views']/i"), "'Views' option not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Save Options']/i"), "'Search' option not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Export']/i"), "'Export' option not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-sidebar//div[@class='iscsidebar-container']"), "'Navigation' Sidebar not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='DataImage']/div[not(contains(@class, 'hide'))]"), "'Products' not present.");
            if (driver._isElementPresent("id", "borderLayout_eGridPanel"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid]"), "Header Row not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-body-container']//div[@row]"), "Rows not present.");
                Results.WriteStatus(test, "Pass", "Verified, Records are displayed in Table.");
            }
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='imageView']/div"), "Records not present.");
                Results.WriteStatus(test, "Pass", "Verified, Records are displayed in Tiles.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'pagination')]//li/a"), "'Pagination' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'pagination')]//button"), "'Page Content View' not present.");

            Results.WriteStatus(test, "Pass", "Verified, Home Page In Detail.");
            return new Home(driver, test);
        }

        ///<summary>
        ///Verify Alert Popup Message and click button
        ///</summary>
        ///<returns></returns>
        public Home VerifyAlertPopupMessageAndClickButton(string message, string buttonName, int waitToClick = 1)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Alert Popup not present.");
            IList<IWebElement> popupCollection = driver._findElements("xpath", "//div[@class='modal-content']");
            int currPopupNum = popupCollection.Count - 1;

            IList<IWebElement> alertHeaderColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//h4");
            Assert.AreNotEqual(0, alertHeaderColl.Count, "'Alert' header not present.");

            IList<IWebElement> alertTextColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//div[contains(text(), '" + message + "')]");
            Assert.AreNotEqual(0, alertTextColl.Count, "'Alert' text message not present.");

            IList<IWebElement> alertButtonColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//button[text()]");
            Assert.AreNotEqual(0, alertButtonColl.Count, "'Alert' Buttons not present.");

            Thread.Sleep(waitToClick * 1000);
            bool avail = false;
            foreach (IWebElement button in alertButtonColl)
                if (button.Text.ToLower().Contains(buttonName.ToLower()))
                {
                    avail = true;
                    driver._clickByJavaScriptExecutor("//div[@class='modal-content']//*[contains(text(), '" + buttonName + "')]");
                    //button.Click();
                    break;
                }
            Assert.IsTrue(avail, "'" + buttonName + "' not found on alert popup.");

            Results.WriteStatus(test, "Pass", "Verified, Alert Popup Message '" + message + "' and clicked '" + buttonName + "' button.");
            return new Home(driver, test);
        }

        ///<summary>
        ///Compare String Lists
        ///</summary>
        ///<param name="list1">First String List</param>
        ///<param name="list2">Second String List</param>
        ///<param name="inOdrer">Whether both string lists are to be compared in order</param>
        ///<param name="match">Whether both string lists should match</param>
        ///<returns></returns>
        public Home CompareStringLists(string[] list1, string[] list2, bool inOdrer = true, bool match = true)
        {
            if (match)
            {
                Assert.AreEqual(list1.Length, list2.Length, "Given String lists don't match as both of them are of different lengths.");

                if (inOdrer)
                {
                    Assert.IsTrue(list1.SequenceEqual(list2), "Given String lists don't match.");
                    Results.WriteStatus(test, "Pass", "Verified, Given String lists match in order as expected.");
                }
                else
                {
                    foreach (string list1Item in list1)
                    {
                        bool avail = true;
                        foreach (string list2Item in list2)
                            if (list1Item.Equals(list2Item))
                            {
                                avail = true;
                                break;
                            }
                        Assert.IsTrue(avail, "'" + list1Item + "' of list1 not found in list2.");
                    }
                    Results.WriteStatus(test, "Pass", "Verified, Given String lists match as expected.");
                }
            }
            else
            {
                if (list1.Length != list2.Length)
                    Results.WriteStatus(test, "Pass", "Verified, Given String lists don't match as both of them are of different lengths.");
                else
                {
                    if (inOdrer)
                        Assert.IsFalse(list1.SequenceEqual(list2), "Given String lists match.");
                    else
                    {
                        int matchIndex = 0;
                        foreach (string list1Item in list1)
                        {
                            foreach (string list2Item in list2)
                                if (list1Item.Equals(list2Item))
                                {
                                    ++matchIndex;
                                    break;
                                }
                        }
                        Assert.AreNotEqual(matchIndex, list1.Length, "Given String lists match.");
                    }
                    Results.WriteStatus(test, "Pass", "Verified, Given String lists don't match as expected.");
                }
            }

            return new Home(driver, test);
        }

        /// <summary>
        /// Verify File Downloaded Or Not for Chart
        /// </summary>
        /// <param name="fileName">File Name to Verify</param>
        /// <param name="FileType">File Extension to Verify</param>
        /// <returns></returns>
        public string VerifyFileDownloadedOrNotOnScreen(string fileName, string FileType)
        {
            bool Exist = false;
            string FilePath = "";
            string Path = ExtentManager.ResultsDir;
            string[] filePaths = Directory.GetFiles(Path, FileType);

            foreach (string filePath in filePaths)
            {
                FileInfo ThisFile = new FileInfo(filePath);
                if (filePath.Contains(fileName + "-" + DateTime.Today.ToString("yyyyMMdd")) || filePath.Contains(fileName))
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

        ///<summary>
        ///Verify Left Navigation Menu List And Select Option
        ///</summary>
        ///<param name="option">Option to be clicked from list</param>
        ///<param name="toggleExpand">Whether to collapse Navigation Menu</param>
        ///<returns></returns>
        public Home VerifyLeftNavigationMenuListAndSelectOption(string option = "", bool toggleExpand = false)
        {
            string[] menuList = new string[] { "SEARCH", "Promo Search", "Saved Searches", "REPORTING",
                "Summary", "Calendar", "Retailer Activity", "Category & Brand Share", "Pricing & Promotions",
                "FlashReports", "Numerator Labs", "My Reports", "Settings", "Contact Us", "Debug Info."};

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='iscsidebar-container']"), "Left Navigation Menu not present");
            Assert.IsTrue(driver._waitForElement("xpath", "//img[@src='/Images/isc-NumeratorLogo.png']", 15), "Numerator Logo not found on Home Page.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='iscsidebar-container']//li"), "Left Navigation Menu List not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='iscsidebar-container']//div[@class='iscopenbtn']"), "Left Navigation Menu Collapse/Expand button not present");

            IList<IWebElement> menuItemColl = driver._findElements("xpath", "//div[@class='iscsidebar-container']//li");

            IWebElement optionButton = null;

            foreach (string navItemName in menuList)
            {
                bool avail = false;
                foreach (IWebElement menuItemEle in menuItemColl)
                    if (menuItemEle.Text.ToLower().Contains(navItemName.ToLower()))
                    {
                        avail = true;
                        if (navItemName.ToLower().Equals(option.ToLower()))
                            optionButton = menuItemEle;
                        break;
                    }
                if (!avail)
                    Results.WriteStatus(test, "Info", "'" + navItemName + "' not found in Left Navigation Menu.");
            }

            if (option != "")
            {
                Assert.AreNotEqual(null, optionButton, "'" + option + "' button not present.");
                Thread.Sleep(2000);
                optionButton.Click();
                Thread.Sleep(2000);
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='iscsidebar-container']//li[contains(@class, 'active')]//*[contains(text(),'" + option + "')]"), "'" + option + "' did not get selected.");
            }

            if (toggleExpand)
            {
                driver._click("xpath", "//div[@class='iscsidebar-container']//div[@class='iscopenbtn']");
                Assert.IsTrue(driver._waitForElement("xpath", "//img[@src='/Images/isc-NumeratorSmall.png']"), "Navigation Menu did not collapse.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Left Navigation Menu List And Selected '" + option + "' Option");
            return new Home(driver, test);
        }

        ///<summary>
        ///Get Active Screen Name From Left Navigation Menu
        ///</summary>
        ///<returns></returns>
        public string GetActiveScreenNameFromLeftNavigationMenu()
        {
            string activeScreenName = "";

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='iscsidebar-container']"), "Left Navigation Menu not present");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='iscsidebar-container']//li[contains(@class, 'active')]"), "Active Screen's Name is not highlighted.");
            activeScreenName = driver._getText("xpath", "//div[@class='iscsidebar-container']//li[contains(@class, 'active')]//a");

            Results.WriteStatus(test, "Pass", "Captured, Active Screen is '" + activeScreenName + "'.");
            return activeScreenName;
        }

        ///<summary>
        ///Verify Client and Change If It Does Not Match
        ///</summary>
        ///<returns></returns>
        public Home VerifyClientAndChangeIfItDoesNotMatch(string clientName)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//h1[text()='Settings']"), "Settings page not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[text()='CLIENT ']"), "Client Field Label not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//button[@ng-click='clientClick()']"), "Client Field not present.");
            if (driver._getText("xpath", "//button[@ng-click='clientClick()']/span[text()]").ToLower().Contains(clientName.ToLower()))
            {
                //driver._click("xpath", "//div[contains(@class,'userProfile')]//button[@type='button' and @ng-click='CancelUpdateProfile();']"); Thread.Sleep(1000);
                Results.WriteStatus(test, "Pass", "'" + clientName + "' Client Already Selected.");
            }

            else
            {
                driver._click("xpath", "//button[@ng-click='clientClick()']");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Switch Client Popup not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']//h4"), "Switch Client Popup header not present.");
                Assert.IsTrue(driver._getText("xpath", "//div[@class='modal-content']//h4").ToLower().Contains("switch client"), "Switch Client Popup header text does not match.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//input"), "Searchbox not present.");
                driver._type("xpath", "//div[@class='modal-content']//input", clientName);

                Thread.Sleep(1000);

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='panel-body']//span"), "Client Names not present.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='panel-body']//span[contains(text(), '" + clientName + "')]"), "'" + clientName + "' Client not present.");
                driver._click("xpath", "//div[@class='panel-body']//span[contains(text(), '" + clientName + "')]");
                Thread.Sleep(3000);
                VerifyHomePage();
                VerifyLeftNavigationMenuListAndSelectOption("Settings");
                VerifyClientAndChangeIfItDoesNotMatch(clientName);
                Results.WriteStatus(test, "Pass", "Verified, Client as '" + clientName + "'");
            }
            
            return new Home(driver, test);
        }

        ///<summary>
        ///Verify Email Alert Popup Message and click button
        ///</summary>
        ///<returns></returns>
        public Home VerifyEmailAlertPopupMessageAndClickButton(string message, string buttonName, int waitToClick = 1)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='PopupDiv ui-draggable']"), "Alert Popup not present.");
            IList<IWebElement> popupCollection = driver._findElements("xpath", "//div[@class='PopupDiv ui-draggable']");
            int currPopupNum = popupCollection.Count - 1;

            IList<IWebElement> alertHeaderColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//td[@class='popup-title1']");
            Assert.AreNotEqual(0, alertHeaderColl.Count, "'Alert' header not present.");

            IList<IWebElement> alertTextColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//div[contains(text(), '" + message + "')]");
            Assert.AreNotEqual(0, alertTextColl.Count, "'Alert' text message not present.");

            IList<IWebElement> alertButtonColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//input[@value]");
            Assert.AreNotEqual(0, alertButtonColl.Count, "'Alert' Buttons not present.");

            Thread.Sleep(waitToClick * 1000);
            bool avail = false;
            foreach (IWebElement button in alertButtonColl)
                if (button.GetAttribute("value").ToLower().Contains(buttonName.ToLower()))
                {
                    avail = true;
                    driver._clickByJavaScriptExecutor("//div[@class='PopupDiv ui-draggable']//input[contains(@value, '" + buttonName + "')]");
                    //button.Click();
                    break;
                }
            Assert.IsTrue(avail, "'" + buttonName + "' not found on alert popup.");

            Results.WriteStatus(test, "Pass", "Verified, Email Alert Popup Message '" + message + "' and clicked '" + buttonName + "' button.");
            return new Home(driver, test);
        }

        #endregion

        #region New Methods

        /// <summary>
        /// Verify Alert popup window with message
        /// </summary>
        /// <param name="message"></param>
        /// <param name="buttonName">Alert Message</param>
        /// <param name="successMessage">Successfull Result Message</param>
        /// <returns></returns>
        public Home verifyAlertPopupWindowWithMessage(string message, string buttonName, string successMessage)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Alert Popup not present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//h4[text()='Alert']"), "'Alert' header not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//div[@ng-bind-html]"), "Alert Message not Present.");
            Assert.AreEqual(message, driver._getText("xpath", "//div[@class='modal-content']//div[@ng-bind-html]"), "'" + message + "' Alert Message not match'");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//button[@id='btnClose']"), "'Close' icon not present on Popup window.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//button[@type='button' and text()='Ok']"), "'Ok' Button not present on Popup window.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//button[@type='button' and text()='Cancel']"), "'Cancel' Button not present on Popup window.");

            driver._clickByJavaScriptExecutor("//div[@class='modal-content']//button[@type='button' and text()='" + buttonName + "']");

            if (buttonName == "Ok")
                verifySuccessAlertPopupWindowWithMessage(successMessage);

            Results.WriteStatus(test, "Pass", "Verified, Alert Popup Message '" + message + "' and clicked '" + buttonName + "' button.");
            return new Home(driver, test);
        }

        /// <summary>
        /// Verify Successfull Alert Popup window with message
        /// </summary>
        /// <param name="message">Message to verify</param>
        /// <returns></returns>
        public Home verifySuccessAlertPopupWindowWithMessage(string message)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Alert Popup not present.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//h4[contains(text(),'Alert')]"), "'Alert' header not present.");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//div[@ng-bind-html]"), "Alert Message not Present.");
            Assert.AreEqual(message, driver._getText("xpath", "//div[@class='modal-content']//div[@ng-bind-html]"), "'" + message + "' Alert Message not match'");

            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//button[@id='close']"), "'Close' icon not present on Popup window.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='modal-content']//button[@type='button' and text()='Okay, Got It']"), "'Okay, Got It' icon not present on Popup window.");

            driver._clickByJavaScriptExecutor("//div[@class='modal-content']//button[@type='button' and text()='Okay, Got It']");
            Thread.Sleep(1000);

            Results.WriteStatus(test, "Pass", "Verified '" + message + "' Message and clicked 'Okay, Got It' Button.");
            return new Home(driver, test);
        }

        /// <summary>
        /// Verify Pricing & Promotions screen
        /// </summary>
        /// <returns></returns>
        public Home verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen(string navigationTitle, string subMenuName)
        {
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@ng-show='showSubMenu']//div[@class='header-text']/h1[text()='" + navigationTitle + "']"), "'" + navigationTitle + "' Header name not match.");
            Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='report-item enabled']"), "Report Item's not present.");
            IList<IWebElement> titlesList = driver._findElements("xpath", "//div[@class='report-item enabled']");

            #region Datasheet

            string dataFromSheet = Common.DirectoryPath + ConfigurationManager.AppSettings["DataSheetDir"] + "\\Sub Menu Screens.xlsx";
            string[] Media = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "report media", navigationTitle);
            string[] Title = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "title", navigationTitle);
            string[] Desc = Spreadsheet.GetMultipleValueOfField(dataFromSheet, "desc", navigationTitle);

            #endregion

            for (int i = 0; i < titlesList.Count; i++)
            {
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='report-item enabled']//div[@class='report-media']/i[@class='" + Media[0].ToString() + "']"), "Media Type (Icon) not match for '" + Title[0].ToString() + "'.");
                Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='report-item enabled']//div[@class='title cursorpointer']/a[text()='" + Title[0].ToString() + "']"), "'" + Title[0].ToString() + "' Title not match.");
                //if (Desc[0].ToString() != "-")
                //    Assert.AreEqual(true, driver._isElementPresent("xpath", "//div[@class='report-item enabled']//div[@class='report-content']/div[contains(@class,'description') and contains(text(),'" + Desc[0].ToString() + "')]"), "Description not match of '" + Title[0].ToString() + "'.");
            }

            driver._clickByJavaScriptExecutor("//div[@class='report-item enabled']//div[@class='title cursorpointer']/a[text()='" + subMenuName + "']");
            Thread.Sleep(1000);
            driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
            driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");
            Results.WriteStatus(test, "Pass", "Verified Pricing & Promotions Screen with Title and Descriptions and click '" + subMenuName + "' screen.");
            return new Home(driver, test);
        }

        #endregion

        #region Medlib Methods

        /// <summary>
        /// Verify and Edit More options in Search Criteria
        /// </summary>
        /// <param name="optionName">Option Name</param>
        /// <returns></returns>
        public Home VerifyAndEditMoreOptionsInSearchCriteria(string optionName = "")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]"), "More Options not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/i"), "More Options Expand/Contract arrow not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/span"), "No. of More Options selected not present.");
            string previousNumber = driver._getText("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/span");
            int prevNum = 0;
            previousNumber = previousNumber.Replace(" ", "").Replace(")", "").Replace("(", "");
            Assert.IsTrue(int.TryParse(previousNumber, out prevNum), "Couldn't convert '" + previousNumber + "' to int");

            if (driver._getAttributeValue("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/i", "class").Contains("down"))
                driver._click("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/i[contains(@class, 'up')]"), "More Options not expanded.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'filter-block')]//div[@class='filter-label']/div"), "Option Names not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'filter-block')]//div[@class='prompt-answer-container']"), "Option Containers not present");

            IList<IWebElement> optionNameColl = driver._findElements("xpath", "//div[contains(@class, 'filter-block')]//div[@class='filter-label']/div");
            IList<IWebElement> optionFieldColl = driver._findElements("xpath", "//div[contains(@class, 'filter-block')]//div[@class='prompt-answer-container']");

            if (optionName == "" || optionName.ToLower().Equals("random"))
            {
                Random rand = new Random();
                int x = rand.Next(0, optionNameColl.Count);
                optionName = optionNameColl[x].Text;
            }

            bool avail = false;
            for (int i = 0; i < optionNameColl.Count; i++)
                if (optionNameColl[i].Text.ToLower().Equals(optionName.ToLower()))
                {
                    avail = true;
                    IList<IWebElement> displayedTextColl = optionFieldColl[i]._findElementsWithinElement("xpath", ".//span");
                    if (!(displayedTextColl[0].Text.ToLower().Contains("any")
                        || displayedTextColl[0].Text.ToLower().Contains("include")
                        || displayedTextColl[0].Text.ToLower().Contains("exclude")
                        || displayedTextColl[0].Text.ToLower().Contains("none")))
                    {
                        IList<IWebElement> crossColl = optionFieldColl[i]._findElementsWithinElement("xpath", ".//i[contains(@class, 'remove')]");
                        Assert.AreNotEqual(0, crossColl.Count, "Cross icon not present to remove selection.");
                        crossColl[0].Click();
                        Thread.Sleep(1000);
                        displayedTextColl = optionFieldColl[i]._findElementsWithinElement("xpath", ".//span");
                        Assert.IsTrue(displayedTextColl[0].Text.ToLower().Contains("any"), "Previous selection not removed from '" + optionName + "'");
                    }

                    displayedTextColl[0].Click();
                    Thread.Sleep(2000);
                    if (optionName.Contains(":"))
                        optionName = optionName.Substring(0, optionName.Length - 1);
                    Search searchPage = new Search(driver, test);
                    searchPage.VerifySearchPageAndSelectCategory(optionName, null, "", "Run Report");
                    Thread.Sleep(2000);
                    string newNumber = driver._getText("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/span");
                    int newNum = 0;
                    newNumber = newNumber.Replace(" ", "").Replace(")", "").Replace("(", "");
                    Assert.IsTrue(int.TryParse(newNumber, out newNum), "Couldn't convert '" + newNumber + "' to int");
                    Assert.AreNotEqual(newNum, prevNum, "Selected More Options number did not change");
                }
            Assert.IsTrue(avail, "'" + optionName + "' not found.");

            Results.WriteStatus(test, "Pass", "Verified, And Edited '" + optionName + "' In Search Criteria.");
            return new Home(driver, test);
        }

        #endregion

        #region Vishal

        /// <summary>
        /// Select option from Left Navigation Menu list
        /// </summary>
        /// <param name="option">Option Name</param>
        /// <returns></returns>
        public Home SelectOptionFromLeftNavigationMenuList(string option)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='iscsidebar-container']"), "Left Navigation Menu not present");
            Assert.IsTrue(driver._waitForElement("xpath", "//img[@src='/Images/isc-NumeratorLogo.png']", 15), "Numerator Logo not found on Home Page.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='iscsidebar-container']//li"), "Left Navigation Menu List not present");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='iscsidebar-container']//div[@class='iscopenbtn']"), "Left Navigation Menu Collapse/Expand button not present");

            IList<IWebElement> menuItemColl = driver._findElements("xpath", "//div[@class='iscsidebar-container']//li");
            bool avail = false;
            foreach (IWebElement menuItemEle in menuItemColl)
                if (menuItemEle.Text.ToLower().Contains(option.ToLower()))
                {
                    avail = true;
                    menuItemEle.Click(); Thread.Sleep(2000);
                    break;
                }

            Assert.AreEqual(true, avail, "'" + option + "' Option not present on list.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='iscsidebar-container']//li[contains(@class, 'active')]//*[contains(text(),'" + option + "')]"), "'" + option + "' did not get selected.");

            Results.WriteStatus(test, "Pass", "Verified, Left Navigation Menu List And Selected '" + option + "' Option");
            return new Home(driver, test);
        }

        /// <summary>
        /// Verify Home Page & select Client from List
        /// </summary>
        /// <param name="clientName">Client Name to Select</param>
        /// <returns></returns>
        public Home verifyHomePageAndSelectClientFromList(string clientName = "Procter & Gamble")
        {
            VerifyHomePage();
            VerifyLeftNavigationMenuListAndSelectOption("Settings");
            VerifyClientAndChangeIfItDoesNotMatch(clientName);

            return new Home(driver, test);
        }

        /// <summary>
        /// Click Tab from Promo Search Section
        /// </summary>
        /// <param name="tabName">Tab Name to click</param>
        /// <returns></returns>
        public Home clickTabFromPromoSeachSection(string tabName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='dvHeaderStatistic80']/ul/li"), "'Tabs' not Present on Promo Search Section.");
            IList<IWebElement> tabCollections = driver.FindElements(By.XPath("//div[@id='dvHeaderStatistic80']/ul/li"));
            bool avail = false;

            for (int i = 0; i < tabCollections.Count; i++)
            {
                if (tabCollections[i].Text.Contains(tabName))
                {
                    driver._clickByJavaScriptExecutor("//div[@id='dvHeaderStatistic80']/ul/li[" + (i + 1) + "]/a");
                    driver._waitForElementToBeHidden("xpath", "//div[@style='display: block;']//div[@class='ProcessLoader']");
                    driver._waitForElementToBePopulated("xpath", "//div[@style='display: none;']//div[@class='ProcessLoader']");
                    Assert.IsTrue(tabCollections[i].GetAttribute("class").Contains("active"), "'" + tabName + "' Tab not Selected.");
                    avail = true;
                    break;
                }
            }

            Assert.IsTrue(avail, "'" + tabName + "' Tab not Present on Promo Search Section.");
            Results.WriteStatus(test, "Pass", "Clicked '" + tabName + "' Tab from Promo Search Section.");
            return new Home(driver, test);
        }

        #endregion

    }
}
