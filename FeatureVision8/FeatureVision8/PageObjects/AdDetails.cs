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
    public class AdDetails
    {
        #region Private Variables

        private IWebDriver adDetails;
        private ExtentTest test;
        Home homePage;

        #endregion

        #region Public Methods

        public AdDetails(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.adDetails = driver;
            test = testReturn;
            homePage = new Home(driver, test);

        }

        public IWebDriver driver
        {
            get { return this.adDetails; }
            set { this.adDetails = value; }
        }

        ///<summary>
        ///Select Tab In Madlib Search View
        ///</summary>
        ///<param name="tab">Tab to be selected</param>
        ///<returns></returns>
        public AdDetails SelectTabInMadlibSearchView(string tab)
        {
            Thread.Sleep(5000);
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@ui-view='Header']//li"), "Tabs not present.");
            string[] tabNameList = new string[] { "Promoted Products", "Ad Blocks", "Pages", "Ads" };
            IList<IWebElement> tabCollection = driver._findElements("xpath", "//div[@ui-view='Header']//li");
            IWebElement selectTabEle = null;
            foreach (string tabName in tabNameList)
            {
                bool avail = false;
                foreach (IWebElement tabEle in tabCollection)
                    if (tabEle.Text.ToLower().Contains(tabName.ToLower()))
                    {
                        avail = true;
                        if (tabName.ToLower().Equals(tab.ToLower()))
                            selectTabEle = tabEle;
                        break;
                    }
                Assert.IsTrue(avail, "'" + tabName + "' tab not found.");
            }

            Assert.AreNotEqual(null, selectTabEle, "'" + tab + "' Tab not found.");
            selectTabEle.Click();
            Thread.Sleep(1000);

            Assert.IsTrue(driver._getText("xpath", "//div[@ui-view='Header']//li[contains(@class, 'active')]//p").ToLower().Contains(tab.ToLower()), "'" + tab + "' tab was not selected.");

            Results.WriteStatus(test, "Pass", "Selected, '" + tab + "' Tab In Madlib Search View");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Select Menu Option In Madlib Search View
        ///</summary>
        ///<param name="option">option to be selected</param>
        ///<param name="subOption">Sub Option to be selected</param>
        ///<returns></returns>
        public AdDetails SelectMenuOptionInMadlibSearchView(string option, string subOption = "")
        {
            string[] subOptionNameList = new string[1];
            driver._waitForElementToBeHidden("xpath", "//div[@class='ProcessLoader']");
            string clickAbleoption = "Multi-select";
            if (option.ToLower().Contains("multi"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Multi-select']/i"), "'Multiselect' option not present.");
                string[] specSubOptionNameList = new string[] { "MULTI-SELECT OPTIONS", "Select Visible Ads", "Deselect Visible Ads", "Select All Ads", "Deselect All Ads" };
                Array.Resize(ref subOptionNameList, specSubOptionNameList.Length);
                Array.Copy(specSubOptionNameList, subOptionNameList, specSubOptionNameList.Length);

                if (driver._isElementPresent("xpath", "//navigation-menu//div[@class='tool ng-scope open']//a[@title='" + clickAbleoption + "']/i") == false)
                    driver._clickByJavaScriptExecutor("//navigation-menu//a[@title='" + clickAbleoption + "']/i");
            }
            else if (option.ToLower().Contains("views"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Views']/i"), "'Views' option not present.");
                string[] specSubOptionNameList = new string[] { "VIEWS", "Tiles", "Table", "VIEW OPTIONS", "Customize Columns", "Sort By", "Show All Ads", "Show Selected Ads" };
                Array.Resize(ref subOptionNameList, specSubOptionNameList.Length);
                Array.Copy(specSubOptionNameList, subOptionNameList, specSubOptionNameList.Length);
                clickAbleoption = "Views";

                if (driver._isElementPresent("xpath", "//navigation-menu//div[@class='tool ng-scope open']//a[@title='" + clickAbleoption + "']/i") == false)
                    driver._clickByJavaScriptExecutor("//navigation-menu//a[@title='" + clickAbleoption + "']/i");
            }
            else if (option.ToLower().Contains("search"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Save Options']/i"), "'Search' option not present.");
                string[] specSubOptionNameList = new string[] { "Save Search" };
                Array.Resize(ref subOptionNameList, specSubOptionNameList.Length);
                Array.Copy(specSubOptionNameList, subOptionNameList, specSubOptionNameList.Length);
                clickAbleoption = "Save Options";

                if (driver._isElementPresent("xpath", "//navigation-menu//div[@class='tool ng-scope open']//a[@title='" + clickAbleoption + "']/i") == false)
                    driver._clickByJavaScriptExecutor("//navigation-menu//a[@title='" + clickAbleoption + "']/i");
            }
            else if (option.ToLower().Contains("export"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//navigation-menu//a[@title='Export']/i"), "'Export' option not present.");
                string[] specSubOptionNameList = new string[] { "EXPORT FORMATS", "Excel", "PDF", "Power Point", "Email Option", "SAVE AS OPTIONS", "Save Recordset", "Reset All Selections" };
                Array.Resize(ref subOptionNameList, specSubOptionNameList.Length);
                Array.Copy(specSubOptionNameList, subOptionNameList, specSubOptionNameList.Length);
                clickAbleoption = "Export";

                if (driver._isElementPresent("xpath", "//navigation-menu//div[@class='tool ng-scope open']//a[@title='" + clickAbleoption + "']/i") == false)
                    driver._clickByJavaScriptExecutor("//navigation-menu//a[@title='" + clickAbleoption + "']/i");
            }

            Thread.Sleep(2000);
            Assert.IsTrue(driver._waitForElement("xpath", "//navigation-menu//div[contains(@class, 'open')]//li"), "'" + option + "' DDL not present.");
            IList<IWebElement> subOptionCollection = driver._findElements("xpath", "//navigation-menu//div[contains(@class, 'open')]//li");
            IWebElement subOptionEle = null;
            foreach (string subName in subOptionNameList)
            {
                bool avail = false;
                foreach (IWebElement subEle in subOptionCollection)
                    if (subEle.Text.ToLower().Contains(subName.ToLower()))
                    {
                        avail = true;
                        if (subName.ToLower().Equals(subOption.ToLower()))
                            subOptionEle = subEle;
                        break;
                    }
                Assert.IsTrue(avail, "'" + subName + "' tab not found.");
            }

            if (subOption != "")
            {
                Assert.AreNotEqual(null, subOptionEle, "'" + subOption + "' Sub Option not found.");
                if (!subOptionEle.GetAttribute("class").Contains("active"))
                {
                    subOptionEle.Click();
                    Thread.Sleep(1000);

                    if (driver._isElementPresent("xpath", "//navigation-menu//div[@class='tool ng-scope open']//a[@title='" + clickAbleoption + "']/i") == true)
                        driver._clickByJavaScriptExecutor("//navigation-menu//a[@title='" + clickAbleoption + "']/i");

                    Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//navigation-menu//div[contains(@class, 'open')]//li"), "'" + option + "' DDL still present.");
                }
            }

            Results.WriteStatus(test, "Pass", "Selected, '" + subOption + "' Option from '" + option + "' Menu List.");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Selection of Records
        ///</summary>
        ///<param name="selected">Whether to verify records as selected or deselected</param>
        ///<param name="visible">Whether to verify only visible records or all records</param>
        ///<returns></returns>
        public AdDetails VerifySelectionOfRecords(bool selected, bool visible = true)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='DataImage']/div[not(contains(@class, 'hide'))]"), "'Products' not present.");
            int recordsPerPage = 0;

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'pagination')]//button[contains(@class, 'active')]/a"), "Records per page selected option not present.");
            string itemsPerPage = driver._getText("xpath", "//div[contains(@class, 'pagination')]//button[contains(@class, 'active')]/a");
            Assert.IsTrue(int.TryParse(itemsPerPage, out recordsPerPage), recordsPerPage + " could not be converted to int.");

            bool nextPage = false;
            if (!visible)
                nextPage = true;

            if (driver._isElementPresent("id", "borderLayout_eGridPanel"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-pinned-left-cols-container']/div[@row]"), "Checkbox for each row not present.");
                do
                {
                    for (int i = 0; i < recordsPerPage; i++)
                    {
                        if (driver._isElementPresent("xpath", "//div[@class='ag-pinned-left-cols-container']/div[@row='" + i + "']"))
                        {
                            driver._scrollintoViewElement("xpath", "//div[@class='ag-pinned-left-cols-container']/div[@row='" + i + "']");
                            if (selected)
                                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@class='ag-pinned-left-cols-container']/div[@row='" + i + "']", "class").Contains("selected"), "Row '" + (i + 1) + "' not selected.");
                            else
                                Assert.IsFalse(driver._getAttributeValue("xpath", "//div[@class='ag-pinned-left-cols-container']/div[@row='" + i + "']", "class").Contains("selected"), "Row '" + (i + 1) + "' is selected.");
                        }
                    }

                    if (!driver._getAttributeValue("xpath", "//div[contains(@class, 'pagination')]//li[contains(@class, 'next')]", "class").Contains("disabled"))
                    {
                        if (nextPage)
                        {
                            driver._click("xpath", "//div[contains(@class, 'pagination')]//li[contains(@class, 'next')]/a");
                            homePage.VerifyHomePage();
                        }
                    }
                    else
                        nextPage = false;

                } while (!visible && nextPage);
            }
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='imageView']/div"), "Records not present.");

                do
                {
                    IList<IWebElement> recordsCollection = driver._findElements("xpath", "//div[@id='imageView']/div");
                    for (int i = 0; i < recordsPerPage; i++)
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", recordsCollection[i]);
                        IList<IWebElement> checkBoxEle = recordsCollection[i]._findElementsWithinElement("xpath", ".//input[@type='checkbox']");
                        if (selected)
                            Assert.AreNotEqual(null, checkBoxEle[0].GetAttribute("checked"), "Record '" + (i + 1) + "' not selected.");
                        else
                            Assert.AreEqual(null, checkBoxEle[0].GetAttribute("checked"), "Record '" + (i + 1) + "' is selected.");
                    }

                    if (!driver._getAttributeValue("xpath", "//div[contains(@class, 'pagination')]//li[contains(@class, 'next')]", "class").Contains("disabled"))
                    {
                        if (nextPage)
                        {
                            driver._click("xpath", "//div[contains(@class, 'pagination')]//li[contains(@class, 'next')]/a");
                            homePage.VerifyHomePage();
                        }
                    }
                    else
                        nextPage = false;

                } while (!visible && nextPage);
            }


            if (selected)
            {
                if (visible)
                    Results.WriteStatus(test, "Pass", "Verified, All visible records are selected.");
                else
                    Results.WriteStatus(test, "Pass", "Verified, All records are selected.");
            }
            else
            {
                if (visible)
                    Results.WriteStatus(test, "Pass", "Verified, All visible records are deselected.");
                else
                    Results.WriteStatus(test, "Pass", "Verified, All records are deselected.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Selection of Records In Madlib Search View");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Customize Your Report Popup
        ///</summary>
        ///<param name="defaultVal">Whether to verify for default value</param>
        ///<param name="popupVisible">Whether Popup is visible</param>
        ///<param name="templateName">Active template name</param>
        ///<returns></returns>
        public AdDetails VerifyCustomizeYourReportPopup(bool popupVisible = true, bool defaultVal = true, string templateName = "Default Template")
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Customize Your Report Popup not present.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-header']/h4"), "Customize Your Report Popup header not present.");
                Assert.AreEqual("Customize your Report", driver._getText("xpath", "//div[@class='modal-header']/h4"), "Customize Your Report Popup header text does not match.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//label[text()='Select template:']"), "'Select Template' label not present in Customize Your Report Popup.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//select/option[@selected='selected']"), "'Select Template' field not present in Customize Your Report Popup.");

                Console.WriteLine("Selected Template Value : " + driver._getText("xpath", "//div[@class='modal-body-filters']//select/option[@selected='selected']"));

                if (defaultVal)
                {
                    if (driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//select/option[@selected='selected' and @label='Default Template']") == false)
                        defaultVal = false;
                    //Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//select/option[@selected='selected' and @label='Default Template']"), "'Select Template' field selected option is not 'Default Template' in Customize Your Report Popup.");
                }
                //Assert.AreEqual("Default Template", driver._getText("xpath", "//div[@class='modal-body-filters']//select/option[@selected='selected']"), "'Select Template' field selected option is not 'Default Template' in Customize Your Report Popup.");
                else
                {
                    string value = driver._getValue("xpath", "//div[@class='modal-body-filters']//select");
                    Assert.AreEqual(templateName, driver._getText("xpath", "//div[@class='modal-body-filters']//select/option[@value='" + value + "']"), "'Select Template' field selected option is not 'Default Template' in Customize Your Report Popup.");
                }

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']//b/button"), "'Drag the columns to right box to be displayed in report:' label not present in Customize Your Report Popup.");
                Assert.AreEqual("Drag the columns to right box to be displayed in report:", driver._getText("xpath", "//div[@class='modal-content']//b/button"), "'Drag the columns to right box to be displayed in report:' label text does not match in Customize Your Report Popup.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//span[text()='All available fields']"), "'All Available Fields' box title not present in Customize Your Report Popup.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//span[text()='Fields to Display']"), "'Fields to Display' box title not present in Customize Your Report Popup.");

                Assert.IsTrue(driver._waitForElement("id", "Master"), "'All Available Fields' box not present in Customize Your Report Popup.");
                Assert.IsTrue(driver._waitForElement("id", "Client"), "'Fields to Display' box not present in Customize Your Report Popup.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='Client']//i[@title='Set To Default' and not(contains(@class, 'hide'))]"), "'Set To Default' icons not present in Fields to Display box in Customize Your Report Popup.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='Client']//i[@title='Edit Field' and not(contains(@class, 'hide'))]"), "'Edit' icons not present in Fields to Display box in Customize Your Report Popup.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//label[text()='Save template as:']"), "'Save template as' label not present in Customize Your Report Popup.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//input[contains(@class, 'inputTextStyle')]"), "'Save template as' field not present in Customize Your Report Popup.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//span[text()='Private']"), "'Private' Radio button label not present.");
                IWebElement radioEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='radio' and @value='0']"));
                Assert.AreNotEqual(null, radioEle, "'Private' Radio button not present.");
                if (defaultVal)
                    Assert.AreNotEqual(null, radioEle.GetAttribute("checked"), "'Private' Radio button not selected by default.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//span[text()='Shared']"), "'Shared' Radio button label not present.");
                radioEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='radio' and @value='1']"));
                Assert.AreNotEqual(null, radioEle, "'Shared' Radio button not present.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//span[text()='Make Active']"), "'Make Active' Checkbox label not present.");
                IWebElement checkBoxEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='checkbox']"));
                Assert.AreNotEqual(null, checkBoxEle, "'Make Active' Checkbox not present.");
                if (defaultVal)
                    Assert.AreNotEqual(null, checkBoxEle.GetAttribute("checked"), "'Make Active' Checkbox not checked by default.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']//div/button[text()]"), "Buttons not present in Customize Your Report Popup");
                string[] buttonNameList = new string[] { "Delete", "Save & Apply", "Apply Selections", "Reset", "Cancel" };
                IList<IWebElement> buttonCollection = driver._findElements("xpath", "//div[@class='modal-content']//div/button[text()]");

                foreach (string buttonName in buttonNameList)
                {
                    bool avail = false;
                    foreach (IWebElement button in buttonCollection)
                        if (button.Text.ToLower().Equals(buttonName.ToLower()))
                        {
                            avail = true;
                            if ((buttonName.ToLower().Equals("delete") || buttonName.ToLower().Equals("reset")) && defaultVal)
                                Assert.IsTrue(button.GetAttribute("class").Contains("disabled"), "'" + buttonName + "' is not disabled by default.");
                            break;
                        }
                    Assert.IsTrue(avail, "'" + buttonName + "' button not found in Customize Your Report Popup");
                }
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "Customize Your Report Popup is present.");
                Results.WriteStatus(test, "Pass", "Verified, Customize Your Report Popup is not present.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Customize Your Report Popup");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Drag Fields In Customize Your Report Popup
        ///</summary>
        ///<param name="targetBox">Box to drag the element to/in</param>
        ///<param name="across">To change the box or the order</param>
        ///<returns></returns>
        public AdDetails DragFieldsInCustomizeYourReportPopup(string targetBox, bool across = true)
        {
            Assert.IsTrue(driver._isElementPresent("id", "Master"), "'All Available Fields' box not present in Customize Your Report Popup.");
            Assert.IsTrue(driver._isElementPresent("id", "Client"), "'Fields to Display' box not present in Customize Your Report Popup.");

            IWebElement sourceEle = null;
            IWebElement targetEle = null;

            if (targetBox.ToLower().Contains("available"))
            {
                if (!across)
                {
                    Results.WriteStatus(test, "Info", "Dragging fields to change order in 'All Available Fields' box is not enabled.");
                    return new AdDetails(driver, test);
                }
                else
                {
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='Master']/div"), "'All Available Fields' box fields not present in Customize Your Report Popup.");
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='Client']/div"), "'Fields to Display' box fields not present in Customize Your Report Popup.");

                    IList<IWebElement> targetElementColl = driver._findElements("xpath", "//div[@id='Master']/div");
                    IList<IWebElement> sourceElementColl = driver._findElements("xpath", "//div[@id='Client']/div");

                    Random rand = new Random();
                    int x = rand.Next(0, sourceElementColl.Count);
                    sourceEle = sourceElementColl[x];
                    targetEle = targetElementColl[0];
                }
            }
            else if (targetBox.ToLower().Contains("display"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='Client']/div"), "'Fields to Display' box fields not present in Customize Your Report Popup.");
                IList<IWebElement> targetElementColl = driver._findElements("xpath", "//div[@id='Client']/div");

                if (across)
                {
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='Master']/div"), "'All Available Fields' box fields not present in Customize Your Report Popup.");
                    IList<IWebElement> sourceElementColl = driver._findElements("xpath", "//div[@id='Master']/div");

                    Random rand = new Random();
                    int x = rand.Next(0, sourceElementColl.Count);
                    sourceEle = sourceElementColl[x];
                    targetEle = targetElementColl[2];
                }
                else
                {
                    Random rand = new Random();
                    int x = rand.Next(0, targetElementColl.Count);
                    sourceEle = targetElementColl[x];
                    targetEle = targetElementColl[x + 2];
                }
            }

            Assert.IsTrue((sourceEle != null) && (targetEle != null), "Either Source or Target element to drag field not found.");

            Actions action = new Actions(driver);
            action.DragAndDrop(sourceEle, targetEle).Perform();
            Thread.Sleep(1000);

            Results.WriteStatus(test, "Pass", "Dragged, Fields In Customize Your Report Popup");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Click Button In Customize Your Report Popup
        ///</summary>
        ///<param name="buttonName">Name of button to be clicked</param>
        ///<returns></returns>
        public AdDetails ClickButtonInPopup(string buttonName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div/button[text()]"), "Buttons not present in Customize Your Report Popup");
            IList<IWebElement> buttonCollection = driver._findElements("xpath", "//div[@class='modal-content']//div/button[text()]");

            bool avail = false;
            foreach (IWebElement button in buttonCollection)
                if (button.Text.ToLower().Equals(buttonName.ToLower()))
                {
                    button.Click();
                    avail = true;
                    break;
                }
            Assert.IsTrue(avail, "'" + buttonName + "' button not found in Customize Your Report Popup");

            Results.WriteStatus(test, "Pass", "Clicked, '" + buttonName + "' Button In Customize Your Report Popup");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Select Template In Customize Your Report Popup
        ///</summary>
        ///<param name="templateName">Name of button to be clicked</param>
        ///<returns></returns>
        public AdDetails SelectTempleteInCustomizeYourReportPopup(string templateName = "")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//select"), "'Select Template' field not present in Customize Your Report Popup.");
            driver._click("xpath", "//div[@class='modal-body-filters']//select");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//select/option"), "'Select Template' field DDL not present in Customize Your Report Popup.");
            IList<IWebElement> templateCollection = driver._findElements("xpath", "//div[@class='modal-body-filters']//select/option");

            if (templateName == "" || templateName.ToLower().Equals("random"))
            {
                Random rand = new Random();
                int x = rand.Next(0, templateCollection.Count);
                templateName = templateCollection[x].Text;
                while (templateName.ToLower().Contains("default") || templateName.ToLower().Contains("current"))
                    templateName = templateCollection[++x].Text;
            }

            bool avail = false;
            foreach (IWebElement template in templateCollection)
                if (template.Text.ToLower().Equals(templateName.ToLower()))
                {
                    template.Click();
                    avail = true;
                    break;
                }

            Thread.Sleep(2000);
            Assert.IsTrue(avail, "'" + templateName + "' button not found in Customize Your Report Popup");

            Results.WriteStatus(test, "Pass", "Selected, '" + templateName + "' template In Customize Your Report Popup");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Capture Fields from Customize Your Report Popup
        ///</summary>
        ///<param name="boxName">Name of box from which list of fields is to be read</param>
        ///<return></return>
        public string[] CaptureFieldsFromCustomizeYourReportPopup(string boxName)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='Master']/div//span[not(contains(@class, 'hide'))]"), "'All Available Fields' box fields not present in Customize Your Report Popup.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='Client']/div//span[not(contains(@class, 'hide'))]"), "'Fields to Display' box fields not present in Customize Your Report Popup.");
            string xpath = "";
            int length = 0, index = 0;
            if (boxName.ToLower().Contains("available"))
            {
                boxName = "All Available Fields";
                xpath = "//div[@id='Master']/div//span[not(contains(@class, 'hide'))]";
            }
            else
            {
                boxName = "Fields to Display";
                xpath = "//div[@id='Client']/div//span[not(contains(@class, 'hide'))]";
            }

            IList<IWebElement> fieldNameColl = driver._findElements("xpath", xpath);
            if (boxName.Equals("All Available Fields"))
                length = fieldNameColl.Count;
            else
            {
                length = fieldNameColl.Count - 2;
                index = 2;
            }
            string[] fieldNameList = new string[length];

            for (int i = index, j = 0; j < fieldNameList.Length; j++, i++)
            {
                if (i % 10 == 0 && i != 0)
                    ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementById('Client').scrollTop+=300", "");
                fieldNameList[j] = fieldNameColl[i].Text;
                Console.WriteLine(j + ": " + fieldNameList[j]);
            }

            Results.WriteStatus(test, "Pass", "Captured, '" + boxName + "' fields from In Customize Your Report Popup");
            return fieldNameList;
        }

        ///<summary>
        ///Capture Fields From Madlib Search Table View
        ///</summary>
        ///<returns></returns>
        public string[] CaptureFieldsFromMadlibSearchTableView()
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='DataImage']/div[not(contains(@class, 'hide'))]"), "'Products' not present.");
            string[] fieldNameList = new string[1];
            int i = 1, j = i - 1;
            string temp = "";
            ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft-=10000", "");

            while (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + i + "]//span[@id='agText']"))
            {
                fieldNameList[j] = driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + i + "]//span[@id='agText']");
                Array.Resize(ref fieldNameList, fieldNameList.Length + 1);
                temp = temp + ";" + fieldNameList[j] + ";";
                Console.WriteLine(j + ": " + fieldNameList[j] + " : " + driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + i + "]//span[@id='agText']"));

                if (!driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (i + 1) + "]//span[@id='agText']"))
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft+=1200", "");
                    Thread.Sleep(1000);
                    if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][1]//span[@id='agText']"))
                    {
                        IList<IWebElement> columnColl = driver._findElements("xpath", "//div[@class='ag-header-container']//div[@colid]//span[@id='agText']");
                        int x = 0;
                        while (x < columnColl.Count && temp.Contains(";" + columnColl[x].Text + ";"))
                            ++x;
                        if (x > columnColl.Count - 1)
                            break;
                        else
                            i = x;
                    }
                    else
                        break;
                }

                if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (i + 1) + "]//span[@id='agText']"))
                    ++i;
                ++j;
            }
            Array.Resize(ref fieldNameList, fieldNameList.Length - 1);

            Results.WriteStatus(test, "Pass", "Captured, Fields from Madlib Search Table View");
            return fieldNameList;
        }

        ///<summary>
        ///Enter Template Name and Select Radio Button
        ///</summary>
        ///<param name="makeActive">Whether to make template active</param>
        ///<param name="shared">Whether to make template shared</param>
        ///<param name="templateName">Template to be saved as</param>
        ///<returns></returns>
        public string EnterTemplateNameAndSelectRadioButton(string templateName = "", bool shared = true, bool makeActive = true)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//input[contains(@class, 'inputTextStyle')]"), "'Save template as' field not present in Customize Your Report Popup.");

            if (templateName == "" || templateName.ToLower().Equals("random"))
                templateName = "TestTemplate" + driver._randomString(3, true);

            driver._type("xpath", "//div[@class='modal-body-filters']//input[contains(@class, 'inputTextStyle')]", templateName);

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//span[text()='Private']"), "'Private' Radio button label not present.");
            IWebElement radioEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='radio' and @value='0']"));
            Assert.AreNotEqual(null, radioEle, "'Private' Radio button not present.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//span[text()='Shared']"), "'Shared' Radio button label not present.");
            radioEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='radio' and @value='1']"));
            Assert.AreNotEqual(null, radioEle, "'Shared' Radio button not present.");

            if (shared)
            {
                driver._click("xpath", "//div[@class='modal-body-filters']//span[text()='Shared']");
                Thread.Sleep(1000);
                radioEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='radio' and @value='1']"));
                Assert.AreNotEqual(null, radioEle.GetAttribute("checked"), "'Shared' Radio button did not get selected.");
                Results.WriteStatus(test, "Pass", "Selected, 'Shared' Radio Button.");
            }
            else
            {
                radioEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='radio' and @value='0']"));
                Assert.AreNotEqual(null, radioEle.GetAttribute("checked"), "'Private' Radio button not selected by default.");
                Results.WriteStatus(test, "Pass", "Selected, 'Private' Radio Button.");
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//span[text()='Make Active']"), "'Make Active' Checkbox label not present.");
            IWebElement checkBoxEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='checkbox']"));
            Assert.AreNotEqual(null, checkBoxEle, "'Make Active' Checkbox not present.");

            if (makeActive && checkBoxEle.GetAttribute("checked") == null)
            {
                driver._click("xpath", "//div[@class='modal-body-filters']//span[text()='Make Active']");
                Thread.Sleep(1000);
                checkBoxEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='checkbox']"));
                Assert.AreNotEqual(null, checkBoxEle.GetAttribute("checked"), "'Make Active' Checkbox did not get checked.");
                Results.WriteStatus(test, "Pass", "Checked, 'Make Active' Checkbox.");
            }
            else if (!makeActive && checkBoxEle.GetAttribute("checked") != null)
            {
                checkBoxEle.Click();
                Thread.Sleep(1000);
                checkBoxEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='checkbox']"));
                Assert.AreEqual(null, checkBoxEle.GetAttribute("checked"), "'Make Active' Checkbox did not get checked.");
                Results.WriteStatus(test, "Pass", "Unchecked, 'Make Active' Checkbox.");
            }

            Results.WriteStatus(test, "Pass", "Entered, '" + templateName + "' as Template Name and selected Radio button.");
            return templateName;
        }

        ///<summary>
        ///Edit Field Name From Fields To Display Box In Customize Your Report Popup
        ///</summary>
        ///<param name="fieldName">Field Name to be edited</param>
        ///<param name="setToDefault">To Edit or Set to Default</param>
        ///<returns></returns>
        public string EditFieldNameFromFieldsToDisplayBoxInCustomizeYourReportPopup(string fieldName = "", bool setToDefault = false)
        {

            Assert.IsTrue(driver._isElementPresent("id", "Client"), "'Fields to Display' box not present in Customize Your Report Popup.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='Client']/div//span[not(contains(@class, 'hide'))]"), "'Fields to Display' box fields not present in Customize Your Report Popup.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='Client']//i[@title='Set To Default' and not(contains(@class, 'hide'))]"), "'Set To Default' icons not present in Fields to Display box in Customize Your Report Popup.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='Client']//i[@title='Edit Field' and not(contains(@class, 'hide'))]"), "'Edit' icons not present in Fields to Display box in Customize Your Report Popup.");

            IList<IWebElement> setToDefaultIconColl = driver._findElements("xpath", "//div[@id='Client']//i[@title='Set To Default' and not(contains(@class, 'hide'))]");
            IList<IWebElement> editIconColl = driver._findElements("xpath", "//div[@id='Client']//i[@title='Edit Field' and not(contains(@class, 'hide'))]");
            IList<IWebElement> fieldNameColl = driver._findElements("xpath", "//div[@id='Client']/div//span[not(contains(@class, 'hide'))]");
            int index = -1;
            if (fieldName == "" || fieldName.ToLower().Equals("random"))
            {
                Random rand = new Random();
                index = rand.Next(2, fieldNameColl.Count);
                fieldName = fieldNameColl[index].Text;
            }
            else
            {
                for (int i = 2; i < fieldNameColl.Count; i++)
                    if (fieldNameColl[i].Text.ToLower().Equals(fieldName.ToLower()))
                    {
                        index = i;
                        break;
                    }
                Assert.Greater(index, -1, "'" + fieldName + "' not found in Fields to Display box.");
            }

            if (!setToDefault)
            {
                editIconColl[index - 2].Click();
                IList<IWebElement> fieldNameTextColl = driver._findElements("xpath", "//div[@id='Client']/div//input");
                fieldNameTextColl[index].SendKeys(driver._randomString(4));
                Actions action = new Actions(driver);
                action.MoveToElement(fieldNameColl[index + 1]).Click().Perform();

                fieldNameColl = driver._findElements("xpath", "//div[@id='Client']/div//span[not(contains(@class, 'hide'))]");
                fieldName = fieldNameColl[index].Text;

                Results.WriteStatus(test, "Pass", "Edited, Field Name From Fields To Display Box");
            }
            else
            {
                setToDefaultIconColl[index - 2].Click();
                Results.WriteStatus(test, "Pass", "Set to Default, Field Name From Fields To Display Box");
            }

            return fieldName;
        }

        ///<summary>
        ///Verify Choose Multiple Columns to Sort by Popup
        ///</summary>
        ///<param name="button">Button to be clicked</param>
        ///<param name="popupVisible">Whether popup is visible</param>
        ///<param name="verifyAddOption">Whether to verify Add another sort option button</param>
        ///<returns></returns>
        public AdDetails VerifyChooseMultipleColumnsToSortByPopup(bool popupVisible = true, bool verifyAddOption = false, string button = "")
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Choose Multiple Columns to Sort by Popup not present.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-header']/h4"), "Choose Multiple Columns to Sort by Popup header not present.");
                Assert.AreEqual("Choose Multiple Columns to Sort by...", driver._getText("xpath", "//div[@class='modal-header']/h4"), "Choose Multiple Columns to Sort by Popup header text does not match.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//label[text()='Sort by']"), "'Sort by' field label not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//label[text()='Then sort by']"), "'Then sort by' field label not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//select[@ng-model='option.SelectedField']"), "'Sort by' Text Fields not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//select[@ng-model='option.SelectedSortOrder']"), "'Sort by' Order By not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//a"), "'Add another sort option' button not present.");
                Assert.AreEqual("Add another sort option", driver._getText("xpath", "//div[@class='modal-body-filters']//a"), "'Add another sort option' button text does not match.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'modal-footer')]//button"), "Buttons not present in Choose Multiple Columns to Sort by Popup");
                string[] buttonNameList = new string[] { "Sort Report", "Cancel" };
                IList<IWebElement> buttonCollection = driver._findElements("xpath", "//div[contains(@class, 'modal-footer')]//button");

                IWebElement clickButton = null;

                foreach (string buttonName in buttonNameList)
                {
                    bool avail = false;
                    foreach (IWebElement buttonEle in buttonCollection)
                        if (buttonEle.Text.ToLower().Equals(buttonName.ToLower()))
                        {
                            avail = true;
                            if (buttonName.ToLower().Equals(button.ToLower()))
                                clickButton = buttonEle;
                            break;
                        }
                    Assert.IsTrue(avail, "'" + buttonName + "' button not found in Choose Multiple Columns to Sort by Popup");
                }

                if (verifyAddOption)
                {
                    while (driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//a"))
                    {
                        driver._click("xpath", "//div[@class='modal-body-filters']//a");
                        Thread.Sleep(500);
                    }

                    IList<IWebElement> sortByFieldColl = driver._findElements("xpath", "//div[@class='modal-body-filters']//select[@ng-model='option.SelectedField']");
                    IList<IWebElement> orderByFieldColl = driver._findElements("xpath", "//div[@class='modal-body-filters']//select[@ng-model='option.SelectedSortOrder']");

                    Assert.AreEqual(8, sortByFieldColl.Count, "More or Less than 8 Levels allowed to add for sorting.");
                    Assert.AreEqual(8, orderByFieldColl.Count, "More or Less than 8 Levels allowed to add for sorting.");
                    Results.WriteStatus(test, "Pass", "Verified, Sort by can be added in popup up to 7 level.");

                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//i[not(contains(@class, 'disabled'))]"), "'Remove' Icons not present against Then Sort By Fields.");
                    IList<IWebElement> removeIconColl = driver._findElements("xpath", "//div[@class='modal-body-filters']//i[not(contains(@class, 'disabled'))]");
                    IList<IWebElement> allRemoveIcons = driver._findElements("xpath", "//div[@class='modal-body-filters']//i");

                    while (allRemoveIcons.Count > 3)
                    {
                        removeIconColl[removeIconColl.Count - 1].Click();
                        removeIconColl = driver._findElements("xpath", "//div[@class='modal-body-filters']//i[not(contains(@class, 'disabled'))]");
                        allRemoveIcons = driver._findElements("xpath", "//div[@class='modal-body-filters']//i");
                    }

                    Assert.LessOrEqual(removeIconColl.Count, 3, "More or Less than 3 Sort By Fields present.");
                }

                if (button != "")
                {
                    Assert.AreNotEqual(null, clickButton, "'" + button + "' button not found to be clicked.");
                    clickButton.Click();
                }
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "Choose Multiple Columns to Sort by Popup is present.");
                Results.WriteStatus(test, "Pass", "Verified, Choose Multiple Columns to Sort by Popup is closed.");
            }



            Results.WriteStatus(test, "Pass", "Verified, Choose Multiple Columns to Sort by Popup.");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Capture Sort By Fields In Order
        ///</summary>
        ///<returns></returns>
        public string[] CaptureSortByFieldsInOrder()
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//select[@ng-model='option.SelectedField' and not(@disabled)]"), "'Sort by' Text Fields not present.");
            IList<IWebElement> sortByFieldColl = driver._findElements("xpath", "//div[@class='modal-body-filters']//select[@ng-model='option.SelectedField' and not(@disabled)]");
            string[] sortByFieldsList = new string[sortByFieldColl.Count];

            for (int i = 0; i < sortByFieldsList.Length; i++)
            {
                string value = sortByFieldColl[i].GetAttribute("value");
                IList<IWebElement> optionCollection = sortByFieldColl[i]._findElementsWithinElement("xpath", ".//option");
                bool avail = false;
                foreach (IWebElement option in optionCollection)
                    if (option.GetAttribute("value").Equals(value))
                    {
                        sortByFieldsList[i] = option.Text;
                        avail = true;
                        break;
                    }
                Assert.IsTrue(avail, "Sort By Field No. '" + (i + 1) + "' not captured.");
            }

            Results.WriteStatus(test, "Pass", "Captured, Sort By Fields In Order");
            return sortByFieldsList;
        }

        ///<summary>
        ///Verify Sort By Fields Order in Madlib Search Grid
        ///</summary>
        ///<param name="sortByFieldsList">List to Verify From</param>
        ///<returns></returns>
        public AdDetails VerifySortByFieldsOrderInMadLibSearchGrid(string[] sortByFieldsList)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='DataImage']/div[not(contains(@class, 'hide'))]"), "'Products' not present.");
            string temp = "";

            for (int i = 0; i < sortByFieldsList.Length; i++)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft-=10000", "");
                bool avail = false;
                int j = 1;
                while (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']"))
                {
                    temp = temp + ";" + driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']") + ";";
                    if (driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']").Equals(sortByFieldsList[i]))
                    {
                        avail = true;
                        Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[contains(@id, 'Sort') and text()]"), "Column '" + sortByFieldsList[i] + "' is not sorted.");
                        Assert.IsTrue(driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[contains(@id, 'Sort') and text()]").Contains((i + 1).ToString()), "Sorting Order is not as expected for Column '" + sortByFieldsList[i] + "'.");
                        break;
                    }
                    if (!driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (j + 1) + "]//span[@id='agText']"))
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft+=1600", "");
                        Thread.Sleep(1000);
                        if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][1]//span[@id='agText']"))
                        {
                            IList<IWebElement> columnColl = driver._findElements("xpath", "//div[@class='ag-header-container']//div[@colid]//span[@id='agText']");
                            int x = 0;
                            while (x < columnColl.Count && temp.Contains(";" + columnColl[x].Text + ";"))
                                ++x;
                            if (x > columnColl.Count - 1)
                                break;
                            else
                                j = x;
                        }
                        else
                            break;
                    }

                    if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (j + 1) + "]//span[@id='agText']"))
                        ++j;
                }
                Assert.IsTrue(avail, "'" + sortByFieldsList[i] + "' not found in MadLib Search Result Grid.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Fields are sorted in expected order in Madlib Search Grid");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Select Records From Madlib Search Result Grid
        ///</summary>
        ///<param name="num">Select Or Verify Selection of No. of Rows</param>
        ///<param name="select">Whether to select or not</param>
        ///<returns></returns>
        public AdDetails SelectRecordsFromMadlibSearchResultGrid(int num = 1, bool select = true)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='ag-body-container']/div"), "Rows not present.");

            if (select)
            {
                for (int i = 0; i < num; i++)
                {
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-body-container']/div[@row='" + i + "']"), "Row '" + (i + 1) + "' not present.");
                    driver._scrollintoViewElement("xpath", "//div[@class='ag-body-container']/div[@row='" + i + "']");
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-pinned-left-cols-container']/div[@row='" + i + "']//span[@class='ag-selection-checkbox']"), "Checkbox for Row '" + (i + 1) + "' not present.");
                    driver._click("xpath", "//div[@class='ag-pinned-left-cols-container']/div[@row='" + i + "']//span[@class='ag-selection-checkbox']");
                    Thread.Sleep(1000);
                    Assert.IsTrue(driver._getAttributeValue("xpath", "//div[@class='ag-body-container']/div[@row='" + i + "']", "class").Contains("selected"), "Row '" + (i + 1) + "' did not get selected.");
                }
            }

            driver._scrollintoViewElement("xpath", "//div[@class='ag-body-container']/div[@row='0']");

            int selectedNum = 0, j = 0;

            while (selectedNum < num)
            {
                driver._scrollintoViewElement("xpath", "//div[@class='ag-body-container']/div[@row='" + j + "']");
                if (driver._getAttributeValue("xpath", "//div[@class='ag-body-container']/div[@row='" + j + "']", "class").Contains("selected"))
                    ++selectedNum;
                ++j;
            }

            Assert.AreEqual(selectedNum, num, "'" + num + "' rows were not selected.");

            Results.WriteStatus(test, "Pass", "Selected, '" + num + "' Records From Madlib Search Result Grid");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Show All Ads or Show Selected Option From View Menu
        ///</summary>
        ///<param name="num">Number of selected Ads</param>
        ///<param name="selected">Verify Show Selected Ads Option</param>
        ///<returns></returns>
        public AdDetails VerifyShowAllOrShowSelectedOptionFromViewMenu(bool selected = true, int num = 1)
        {
            if (num == 0 && selected)
            {
                homePage.VerifyAlertPopupMessageAndClickButton("There are no selected records.", "Okay, Got It");
                Results.WriteStatus(test, "Pass", "Verified, Dialog Box stating that no records are selected when no records are selected.");
            }

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='ag-body-container']/div"), "Rows not present.");
            IList<IWebElement> rowsCollection = driver._findElements("xpath", "//div[@class='ag-body-container']/div");

            if (selected)
            {
                Assert.AreEqual(num, rowsCollection.Count, "Either Ads displayed are not all selected or all selected Ads are not displayed.");
                Results.WriteStatus(test, "Pass", "Verified, Only Selected Ads are Dislayed.");
            }
            else
            {
                Assert.Greater(rowsCollection.Count, num, "All Ads are not displayed.");
                Results.WriteStatus(test, "Pass", "Verified, All Ads are Dislayed.");
            }

            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Option for creating your Promoted Products Report Popup
        ///</summary>
        ///<param name="excel">Whether Excel Option is clicked from Export Menu</param>
        ///<param name="popupVisible">Whether popup should be visible</param>
        ///<param name="selectedRows">Whether Ads are selected</param>
        ///<return></return>
        public AdDetails VerifyOptionForCreatingYourPromotedProductsReportPopup(bool popupVisible = true, bool selectedRows = false, bool excel = true)
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "'Option for creating your Promoted Products/Ad Blocks Report' Popup not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//h4"), "'Option for creating your Promoted Products Report' Popup header not present.");
                Assert.AreEqual(true, driver._getText("xpath", "//div[@class='modal-content']//h4").ToLower().Equals("Option for creating your Promoted Products Report".ToLower())
                    || driver._getText("xpath", "//div[@class='modal-content']//h4").ToLower().Equals("Option for creating your Ad Blocks Report".ToLower())
                    || driver._getText("xpath", "//div[@class='modal-content']//h4").ToLower().Equals("Option for creating your Ads Report".ToLower()), "'Option for creating your Promoted Products/Ad Blocks Report' Popup header text does not match.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//span[contains(text(), 'Include links to the Numerator Promotions Intel site.')]"), "'Include links to the Numerator Promotions Intel site.' Label not present");
                IWebElement checkbox = driver.FindElement(By.XPath("//div[@class='modal-content']//input[contains(@ng-model, 'IncludeFeatureVisionHyperlink')]"));
                Assert.AreNotEqual(null, checkbox, "Checkbox for 'Include links to the Numerator Promotions Intel site.' not present");
                Assert.AreNotEqual(null, checkbox.GetAttribute("checked"), "Checkbox for 'Include links to the Numerator Promotions Intel site.' not checked by default.");

                if (selectedRows)
                {
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//label[contains(@ng-if, 'UseOnlySelectedRecods')]/span"), "'Use only the selected records that are checked in the Promoted Products Detail Report.' Label not present");
                    checkbox = driver.FindElement(By.XPath("//div[@class='modal-content']//input[contains(@ng-model, 'UseOnlySelectedRecods')]"));
                    Assert.AreNotEqual(null, checkbox, "Checkbox for 'Use only the selected records that are checked in the Promoted Products Detail Report.' not present");
                }

                if (excel)
                {
                    if (!driver._getText("xpath", "//div[@class='modal-content']//h4").ToLower().Equals("Option for creating your Ads Report".ToLower()))
                    {
                        Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//span[contains(text(), 'Display image in excel comments box.')]"), "'Display image in excel comments box.' Label not present");
                        checkbox = driver.FindElement(By.XPath("//div[@class='modal-content']//input[contains(@ng-model, 'ShowProductImageInExcel')]"));
                        Assert.AreNotEqual(null, checkbox, "Checkbox for 'Display image in excel comments box.' not present");
                    }

                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//span[contains(text(), 'Zip File')]"), "'Zip File' Label not present");
                    checkbox = driver.FindElement(By.XPath("//div[@class='modal-content']//input[contains(@ng-model, 'ZipFile')]"));
                    Assert.AreNotEqual(null, checkbox, "Checkbox for 'Zip File' not present");
                }

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//label[contains(text(), 'Report Name')]"), "'Report Name' Label not present");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//input[@id = 'txtReportName']"), "'Report Name' Text Field not present");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//button[text()= 'Download report']"), "'Download report' Button not present on 'Option for creating your Promoted Products Report' popup");
                if (driver._getText("xpath", "//div[@class='modal-content']//h4").ToLower().Contains("promoted products".ToLower()))
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//button[text()= 'Send mail with attached report']"), "'Send mail with attached report' Button not present on 'Option for creating your Promoted Products Report' popup");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//button[text()= 'Cancel']"), "'Cancel' Button not present on 'Option for creating your Promoted Products Report' popup");
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "'Option for creating your Promoted Products Report' Popup is present.");
                Results.WriteStatus(test, "Pass", "Verified, Option for creating your Promoted Products Report Popup is closed.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Option for creating your Promoted Products Report Popup");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Image Report Options Popup
        ///</summary>
        ///<param name="popupVisible">Whether popup should be visible</param>
        ///<return></return>
        public AdDetails VerifyImageReportOptionsPopup(bool popupVisible = true)
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "'Image Report Options' Popup not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//h4"), "'Image Report Options' Popup header not present.");
                Assert.AreEqual(true, driver._getText("xpath", "//div[@class='modal-content']//h4").Contains("Image Report Options - ") && driver._getText("xpath", "//div[@class='modal-content']//h4").Contains(" pages ") && driver._getText("xpath", "//div[@class='modal-content']//h4").Contains(" currently selected"), "'Image Report Options' Popup header text does not match.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//li/a[text()='Report Templates']"), "'Report Templates' Tab not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//li[contains(@class, 'active')]/a[text()='Report Templates']"), "'Report Templates' Tab is not selected by default.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//li/a[text()='More Options']"), "'More Options' Tab not present.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div[contains(text(), '2 - Select the desired layout option(s), by checking ')]"), "'2 - Select the desired layout option(s)' Label not present");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div[@class='border-box']//div[contains(text(), 'Attachment Max. Size: 10 MB')]"), "'Attachment Max. Size: 10 MB' Label not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//div[@class='border-box']//label[text()= 'Selected Templates']"), "'Selected Templates' Label not present.");

                if (driver._isElementPresent("xpath", "//div[contains(@class,'-box') and not(contains(@class, 'ng-hide'))]//span[contains(text(), 'Use In Report')]"))
                {
                    driver._click("xpath", "//div[contains(@class,'-box') and not(contains(@class, 'ng-hide'))]//span[contains(text(), 'Use In Report')]");
                    Thread.Sleep(1000);
                }

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//button[text()= 'Download Report']"), "'Download report' Button not present on 'Option for creating your Promoted Products Report' popup");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//button[text()= 'Email Report as Attachment']"), "'Email Report as Attachment' Button not present on 'Option for creating your Promoted Products Report' popup");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//button[text()= 'Email Report as Link']"), "'Email Report as Link' Button not present on 'Option for creating your Promoted Products Report' popup");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//button[text()= 'Cancel']"), "'Cancel' Button not present on 'Option for creating your Promoted Products Report' popup");
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "'Image Report Options' Popup is present.");
                Results.WriteStatus(test, "Pass", "Verified, Image Report Options Popup is closed.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Image Report Options Popup");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Report Ready Popup Message and click button
        ///</summary>
        ///<returns></returns>
        public AdDetails VerifyReportReadyPopupMessageAndClickButton(string message, string buttonName)
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Report Ready Popup not present.");
            IList<IWebElement> popupCollection = driver._findElements("xpath", "//div[@class='modal-content']");
            int currPopupNum = popupCollection.Count - 1;

            IList<IWebElement> alertHeaderColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//h4");
            Assert.AreNotEqual(0, alertHeaderColl.Count, "'Report Ready' header not present.");
            Assert.IsTrue(alertHeaderColl[0].Text.Contains(""), "'Report Ready' Popup Header text does not match.");

            IList<IWebElement> alertTextColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//td[contains(text(), '" + message + "')]");
            Assert.AreNotEqual(0, alertTextColl.Count, "'Report Ready' text message not present.");

            IList<IWebElement> alertLinkColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//a");
            Assert.AreNotEqual(0, alertLinkColl.Count, "'Report Ready' Download Link not present.");

            IList<IWebElement> alertButtonColl = popupCollection[currPopupNum]._findElementsWithinElement("xpath", ".//button[text()]");
            Assert.AreNotEqual(0, alertButtonColl.Count, "'Report Ready' Buttons not present.");

            bool avail = false;
            foreach (IWebElement button in alertButtonColl)
                if (button.Text.ToLower().Contains(buttonName.ToLower()))
                {
                    avail = true;
                    button.Click();
                    break;
                }
            Assert.IsTrue(avail, "'" + buttonName + "' not found on Report Ready popup.");

            Results.WriteStatus(test, "Pass", "Verified, Report Ready Popup Message '" + message + "' and clicked '" + buttonName + "' button.");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Get Filter Type For A Column Or Column For A Filter Type
        ///</summary>
        ///<param name="columnName">Column Name to Determine the filter type for</param>
        ///<param name="filterType">Filter type to search column for</param>
        ///<returns></returns>
        public string GetFilterTypeForColumnOrColumnForFilterType(string filterType = "", string columnName = "")
        {
            string columnOrFilterType = "";

            Assert.IsFalse(filterType == "" && columnName == "", "Both Filter Type and Column Name cannot be blank.");
            Assert.IsFalse(filterType != "" && columnName != "", "One of Filter Type and Column Name should be blank.");

            if (filterType != "")
            {
                if (filterType.ToLower().Equals("random"))
                {
                    string[] filterTypeValues = new string[] { "Normal", "Number", "Text" };
                    Random rand = new Random();
                    int x = rand.Next(0, filterTypeValues.Length);
                    filterType = filterTypeValues[x];
                }

               ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft-=10000", "");

                int j = 1;
                string temp = "";
                while (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']"))
                {
                    string thisColumn = driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']");
                    temp = temp + ";" + driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']") + ";";
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agMenu']"), "Filter Icon not present for column '" + thisColumn + "'.");
                    driver._click("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agMenu']");
                    Thread.Sleep(1000);
                    Assert.IsTrue(driver._waitForElement("xpath", "//div[not(@id)]/div[@class='ag-filter']"), "Filter Menu not present for column '" + thisColumn + "'.");

                    if (driver._isElementPresent("xpath", "//input[@placeholder='Search...']") && filterType.ToLower().Equals("normal")
                        || (driver._isElementPresent("xpath", "//input[@placeholder='Numeric value...']") && filterType.ToLower().Equals("number"))
                        || (driver._isElementPresent("xpath", "//input[@placeholder='Filter...']") && filterType.ToLower().Equals("text")))
                    {
                        columnOrFilterType = thisColumn;
                        break;
                    }

                    if (!driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (j + 1) + "]//span[@id='agText']"))
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft+=1600", "");
                        Thread.Sleep(1000);
                        if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][1]//span[@id='agText']"))
                        {
                            IList<IWebElement> columnColl = driver._findElements("xpath", "//div[@class='ag-header-container']//div[@colid]//span[@id='agText']");
                            int x = 0;
                            while (x < columnColl.Count && temp.Contains(";" + columnColl[x].Text + ";"))
                                ++x;
                            if (x > columnColl.Count - 1)
                                break;
                            else
                                j = x;
                        }
                        else
                            break;
                    }

                    if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (j + 1) + "]//span[@id='agText']"))
                        ++j;
                }
                Assert.AreNotEqual("", columnOrFilterType, "No column was found for filter type '" + filterType + "'");
                Results.WriteStatus(test, "Pass", "Column '" + columnOrFilterType + "' has filter type '" + filterType + "'");
            }
            else
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='ag-header-container']//div[@colid]//span[@id='agText']"), "Column Headers are not present in Madlib Search Grid.");
                if (columnName.ToLower().Equals("random"))
                {
                    IList<IWebElement> columnCollection = driver._findElements("xpath", "//div[@class='ag-header-container']//div[@colid]//span[@id='agText']");
                    Random rand = new Random();
                    int x = rand.Next(0, columnCollection.Count);
                    columnName = columnCollection[x].Text;
                }

                ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft-=10000", "");
                string temp = "";
                bool avail = false;
                int j = 1;
                while (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']"))
                {
                    temp = temp + ";" + driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']") + ";";
                    if (driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']").Equals(columnName))
                    {
                        avail = true;
                        Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agMenu']"), "Column '" + columnName + "' does not have filter icon.");
                        driver._click("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agMenu']");
                        Thread.Sleep(1000);
                        Assert.IsTrue(driver._waitForElement("xpath", "//div[not(@id)]/div[@class='ag-filter']"), "Filter Menu not present for column '" + columnName + "'.");

                        if (driver._isElementPresent("xpath", "//input[@placeholder='Search...']"))
                        {
                            columnOrFilterType = "Normal";
                            break;
                        }
                        else if (driver._isElementPresent("xpath", "//input[@placeholder='Numeric value...']"))
                        {
                            columnOrFilterType = "Number";
                            break;
                        }
                        else if (driver._isElementPresent("xpath", "//input[@placeholder='Filter...']"))
                        {
                            columnOrFilterType = "Text";
                            break;
                        }
                    }
                    if (!driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (j + 1) + "]//span[@id='agText']"))
                    {
                        ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft+=1600", "");
                        Thread.Sleep(1000);
                        if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][1]//span[@id='agText']"))
                        {
                            IList<IWebElement> columnColl = driver._findElements("xpath", "//div[@class='ag-header-container']//div[@colid]//span[@id='agText']");
                            int x = 0;
                            while (x < columnColl.Count && temp.Contains(";" + columnColl[x].Text + ";"))
                                ++x;
                            if (x > columnColl.Count - 1)
                                break;
                            else
                                j = x;
                        }
                        else
                            break;
                    }

                    if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (j + 1) + "]//span[@id='agText']"))
                        ++j;
                }
                Assert.IsTrue(avail, "Column '" + columnName + "' was not found in Madlib Search Result Grid.");
                Assert.AreNotEqual("", columnOrFilterType, "No filter type was determined for column '" + columnName + "'");
                Results.WriteStatus(test, "Pass", "Column '" + columnName + "' has filter type '" + columnOrFilterType + "'");
            }

            return columnOrFilterType;
        }

        ///<summary>
        ///Verify Normal Filter Functionality
        ///</summary>
        ///<returns></returns>
        public AdDetails VerifyNormalFilterFunctionality(string column = "", bool removeFilter = true)
        {
            if (column == "" || column.ToLower().Equals("random"))
                column = GetFilterTypeForColumnOrColumnForFilterType("Normal");

            ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft-=10000", "");
            string temp = "";
            bool avail = false;
            int j = 1;
            while (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']"))
            {
                temp = temp + ";" + driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']") + ";";
                if (driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agText']").Equals(column))
                {
                    avail = true;
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agMenu']"), "Column '" + column + "' does not have filter icon.");
                    if (!driver._isElementPresent("xpath", "//div[not(@id)]/div[@class='ag-filter']"))
                        driver._click("xpath", "//div[@class='ag-header-container']//div[@colid][" + j + "]//span[@id='agMenu']");
                    Thread.Sleep(1000);
                    Assert.IsTrue(driver._waitForElement("xpath", "//div[not(@id)]/div[@class='ag-filter']"), "Filter Menu not present for column '" + column + "'.");
                    break;
                }
                if (!driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (j + 1) + "]//span[@id='agText']"))
                {
                    ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementsByClassName('ag-body-viewport customScrollBar')[0].scrollLeft+=1600", "");
                    Thread.Sleep(1000);
                    if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][1]//span[@id='agText']"))
                    {
                        IList<IWebElement> columnColl = driver._findElements("xpath", "//div[@class='ag-header-container']//div[@colid]//span[@id='agText']");
                        int x = 0;
                        while (x < columnColl.Count && temp.Contains(";" + columnColl[x].Text + ";"))
                            ++x;
                        if (x > columnColl.Count - 1)
                            break;
                        else
                            j = x;
                    }
                    else
                        break;
                }

                if (driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (j + 1) + "]//span[@id='agText']"))
                    ++j;
            }
            Assert.IsTrue(avail, "Column '" + column + "' was not found in Madlib Search Result Grid.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//input[@placeholder='Search...']"), "'Search' text field not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//span[text()='(Select All)']"), "'Select All' option not present.");
            IWebElement selectAllCheckbox = driver.FindElement(By.XPath("//input[contains(@id, 'agGridFilterSelectAll')]"));
            Assert.AreNotEqual(null, selectAllCheckbox, "'Select All' switch not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-virtual-list-item']//span"), "'Filter Options' not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//button[contains(@id, 'agGridFilterApplyButton')]"), "'Apply Filter' Button not present.");
            Assert.AreNotEqual(null, driver._getAttributeValue("xpath", "//button[contains(@id, 'agGridFilterApplyButton')]", "disabled"), "'Apply Filter' Button not disabled by default.");

            driver._click("xpath", "//span[text()='(Select All)']");
            Thread.Sleep(1000);

            IList<IWebElement> filterCollection = driver._findElements("xpath", "//div[@class='ag-virtual-list-item']//span");
            foreach (IWebElement filterOption in filterCollection)
                Console.WriteLine(filterOption.GetCssValue("checkbox"));


            //Assert.IsTrue(driver._isElementPresent("xpath", ""), " not present.");
            //Assert.IsTrue(driver._isElementPresent("xpath", ""), " not present.");
            //Assert.IsTrue(driver._isElementPresent("xpath", ""), " not present.");
            //Assert.IsTrue(driver._isElementPresent("xpath", ""), " not present.");
            //Assert.IsTrue(driver._isElementPresent("xpath", ""), " not present.");
            //Assert.IsTrue(driver._isElementPresent("xpath", ""), " not present.");


            Results.WriteStatus(test, "Pass", "Verified, Normal Filter Functionality");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Sorting Functionality In Madlib Search Result Grid
        ///</summary>
        ///<param name="colId">Column Id of Column to be sorted</param>
        ///<returns></returns>
        public int VerifySortingFunctionalityInMadlibSearchResultsGrid(string colId = "1")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid='" + colId + "']"), "Column at Index '" + colId + "' not present.");
            string columnName = driver._getText("xpath", "//div[@class='ag-header-container']//div[@colid='" + colId + "']//span[@id='agText']");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid='" + colId + "']//span[contains(@id, 'Sort') and not(contains(@class, 'hidden'))]"), "'Sorting' Icon at Column at Index '" + colId + "' not present.");

            while (!driver._getAttributeValue("xpath", "//div[@class='ag-header-container']//div[@colid='" + colId + "']//span[contains(@id, 'Sort') and not(contains(@class, 'hidden'))]", "class").Contains("descending")
                && !driver._getAttributeValue("xpath", "//div[@class='ag-header-container']//div[@colid='" + colId + "']//span[contains(@id, 'Sort') and not(contains(@class, 'hidden'))]", "class").Contains("ascending"))
            {
                //Actions action = new Actions(driver);
                //action.MoveToElement(driver.FindElement(By.XPath("//div[@class='ag-header-container']//div[@colid='" + colId + "']//span[@id='agText']"))).Click().Perform();
                driver._clickByJavaScriptExecutor("//div[@class='ag-header-container']//div[@colid='" + colId + "']//span[@id='agText']");
                if (driver._waitForElement("xpath", "//div[@class='modal-content']", 15))
                {
                    homePage.VerifyAlertPopupMessageAndClickButton("Sorting is allowed upto 8 fields only.", "Okay, Got It", 10);
                    return 1;
                }
            }


            bool sortAsc = true;

            if (driver._getAttributeValue("xpath", "//div[@class='ag-header-container']//div[@colid='" + colId + "']//span[contains(@id, 'Sort') and not(contains(@class, 'hidden'))]", "class").Contains("descending"))
                sortAsc = false;

            string[] columnCellValuesList = new string[10];

            for (int i = 0; i < 10; i++)
            {
                driver._scrollintoViewElement("xpath", "//div[@class='ag-body-container']//div[" + (i + 1) + "]/div[@colid='" + colId + "']");
                columnCellValuesList[i] = driver._getText("xpath", "//div[@class='ag-body-container']//div[" + (i + 1) + "]/div[@colid='" + colId + "']");
            }
            driver._scrollintoViewElement("xpath", "//div[@class='ag-body-container']//div[1]/div[@colid='" + colId + "']");

            int x = 0;
            if (int.TryParse(columnCellValuesList[0], out x))
            {
                int[] iColumnCellValuesList = new int[columnCellValuesList.Length];
                for (int i = 0; i < columnCellValuesList.Length; i++)
                {
                    string temp = columnCellValuesList[i];
                    if (temp.Contains("$"))
                        temp = temp.Substring(1);

                    while (temp.Contains(","))
                        temp = temp.Remove(temp.IndexOf(","), 1);

                    Assert.IsTrue(int.TryParse(temp, out iColumnCellValuesList[i]), "Couldn't convert '" + columnCellValuesList[i] + "' to int.");
                }

                int[] origIColumnCellValuesList = new int[iColumnCellValuesList.Length];
                Array.Copy(iColumnCellValuesList, origIColumnCellValuesList, iColumnCellValuesList.Length);
                Array.Sort(iColumnCellValuesList);

                if (sortAsc)
                {
                    Assert.IsTrue(origIColumnCellValuesList.SequenceEqual(iColumnCellValuesList), "'" + columnName + "' not sorted in Ascending order properly.'");
                    Results.WriteStatus(test, "Pass", "'" + columnName + "' sorted in Ascending order successfully.");
                }
                else
                {
                    Array.Reverse(iColumnCellValuesList);
                    Assert.IsTrue(origIColumnCellValuesList.SequenceEqual(iColumnCellValuesList), "'" + columnName + "' not sorted in Descending order properly.'");
                    Results.WriteStatus(test, "Pass", "'" + columnName + "' sorted in Descending order successfully.");
                }
            }
            else if (columnCellValuesList[0].IndexOf('/') == 1 || columnCellValuesList[0].IndexOf('/') == 2)
            {
                DateTime[] dColumnCellValuesList = new DateTime[columnCellValuesList.Length];
                System.Globalization.CultureInfo cultures = new System.Globalization.CultureInfo("en-US");
                for (int i = 0; i < columnCellValuesList.Length; i++)
                    //Assert.IsTrue(DateTime.TryParse(columnCellValuesList[i], "MM/dd/yyyy", System.Globalization.CultureInfo.CurrentCulture, out dColumnCellValuesList[i]), "Couldn't convert '" + columnCellValuesList[i] + "' to Date.");
                    dColumnCellValuesList[i] = Convert.ToDateTime(columnCellValuesList[i], cultures);

                DateTime[] origDColumnCellValuesList = new DateTime[dColumnCellValuesList.Length];
                Array.Copy(dColumnCellValuesList, origDColumnCellValuesList, dColumnCellValuesList.Length);
                Array.Sort(dColumnCellValuesList);

                if (sortAsc)
                {
                    Assert.IsTrue(origDColumnCellValuesList.SequenceEqual(dColumnCellValuesList), "'" + columnName + "' not sorted in Ascending order properly.'");
                    Results.WriteStatus(test, "Pass", "'" + columnName + "' sorted in Ascending order successfully.");
                }
                else
                {
                    Array.Reverse(dColumnCellValuesList);
                    Assert.IsTrue(origDColumnCellValuesList.SequenceEqual(dColumnCellValuesList), "'" + columnName + "' not sorted in Descending order properly.'");
                    Results.WriteStatus(test, "Pass", "'" + columnName + "' sorted in Descending order successfully.");
                }
            }
            else
            {
                string[] origColumnCellValuesList = new string[columnCellValuesList.Length];
                Array.Copy(columnCellValuesList, origColumnCellValuesList, columnCellValuesList.Length);
                Array.Sort(columnCellValuesList);

                if (sortAsc)
                {
                    Assert.IsTrue(origColumnCellValuesList.SequenceEqual(columnCellValuesList), "'" + columnName + "' not sorted in Ascending order properly.'");
                    Results.WriteStatus(test, "Pass", "'" + columnName + "' sorted in Ascending order successfully.");
                }
                else
                {
                    Array.Reverse(columnCellValuesList);
                    Assert.IsTrue(origColumnCellValuesList.SequenceEqual(columnCellValuesList), "'" + columnName + "' not sorted in Descending order properly.'");
                    Results.WriteStatus(test, "Pass", "'" + columnName + "' sorted in Descending order successfully.");
                }
            }

            return 0;
        }

        ///<summary>
        ///Verify Sorting On Multiple Columns
        ///</summary>
        ///<returns></returns>
        public AdDetails VerifySortingOnMultipleColumns(int num = 1)
        {
            ((IJavaScriptExecutor)driver).ExecuteScript("document.body.style.zoom = '0.6'");
            Thread.Sleep(1000);
            for (int i = 0; i < num; i++)
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-header-container']//div[@colid][" + (i + 1) + "]"), "Column at Index '" + (i + 1) + "' not present.");
                string colId = driver._getAttributeValue("xpath", "//div[@class='ag-header-container']//div[@colid][" + (i + 1) + "]", "colid");

                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", driver._findElement("xpath", "//div[@class='ag-header-container']//div[@colid='" + colId + "']"));
                if (i > 7)
                {
                    homePage.VerifyAlertPopupMessageAndClickButton("Sorting is allowed upto 8 fields only.", "Okay, Got It");
                    break;
                }
                Thread.Sleep(1000);
                int retValue = VerifySortingFunctionalityInMadlibSearchResultsGrid(colId);
                if (retValue == 1)
                {
                    Assert.AreEqual(7, i, "Message Displayed on incorrect number of columns");
                    break;
                }
                Thread.Sleep(1000);
            }

            Results.WriteStatus(test, "Pass", "Verified, Sorting On Multiple Columns");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Open View Ad Or Details Popup
        ///</summary>
        ///<param name="openDetails">Whether to open Detail Tab</param>
        ///<returns></returns>
        public AdDetails OpenViewAdOrDetailsPopup(bool openDetails = false)
        {
            if (driver._isElementPresent("id", "borderLayout_eGridPanel"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-body-container']//div[@row]"), "Rows not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='ag-body-container']//div[@row]/div"), "Cells not present.");

                driver._click("xpath", "//div[@class='ag-body-container']//div[@row][1]/div[1]");



                Results.WriteStatus(test, "Pass", "Opened, 'View Ad' Popup from Table View.");
            }
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='imageView']/div"), "Records not present.");
                Actions action = new Actions(driver);
                action.MoveToElement(driver.FindElement(By.XPath("//div[@id='imageView']/div[1]")), 5, 5).Perform();
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='imageView']/div[1]//p/a[contains(text(), 'View Ad')]"), "'View Ad' button not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@id='imageView']/div[1]//p/a[contains(text(), 'Detail')]"), "'Detail' button not present.");

                if (openDetails)
                {
                    driver._click("xpath", "//div[@id='imageView']/div[1]//p/a[contains(text(), 'Detail')]");
                    VerifyDetailPopup();

                    Results.WriteStatus(test, "Pass", "Opened, 'Detail' Popup from Tiles View.");
                }
                else
                {
                    driver._click("xpath", "//div[@id='imageView']/div[1]//p/a[contains(text(), 'View Ad')]");
                    VerifyViewAdPopup();

                    Results.WriteStatus(test, "Pass", "Opened, 'View Ad' Popup from Tiles View.");
                }
            }
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify View Ad Popup
        ///</summary>
        ///<param name="popupVisible">Whether the popup should be visible</param>
        ///<returns></returns>
        public AdDetails VerifyViewAdPopup(bool popupVisible = true)
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "'View Ad' Popup not present.");
                if (!driver._getAttributeValue("xpath", "//div[@class='modal-content']//ul[contains(@class, 'nav')]/li[1]", "class").Contains("active"))
                    driver._click("xpath", "//div[@class='modal-content']//ul[contains(@class, 'nav')]/li[1]");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='mid-size-flyer']//img[@id]"), "Center Image not present in View Ad Popup");
                driver.MouseHoverUsingElement("xpath", "//div[@class='mid-size-flyer']//img[@id]");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='mid-size-flyer']//img[@class='hoverimage']"), "Zoom Icon not present in View Ad Popup");
                Assert.IsTrue(driver._isElementPresent("xpath", "//mt-flyer//div[contains(@class, 'expand')]/div"), "Expand button not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//mt-flyer//div[@class='panel panel-default']"), "Ad Block not present");
                Assert.IsTrue(driver._isElementPresent("xpath", "//mt-flyer//div[contains(@class, 'active')]/ul/li"), "Pages not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//mt-flyer//div[contains(@class, 'active')]/ul/li//span"), "Checkbox on Pages not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//mt-flyer//a[contains(@class, 'carousel-control')]"), "Navigation Arrows on Pages not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-header bgColor']/button"), "'X' button not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-footer']/button"), "'Close' button not present");
                Results.WriteStatus(test, "Pass", "Verified, View Ad Popup.");
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "'View Ad' Popup is still present.");
                Results.WriteStatus(test, "Pass", "Verified, View Ad Popup is closed.");
            }
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Detail Popup
        ///</summary>
        ///<param name="popupVisible">Whether the popup should be visible</param>
        ///<returns></returns>
        public AdDetails VerifyDetailPopup(bool popupVisible = true)
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "'Detail' Popup not present.");
                if (!driver._getAttributeValue("xpath", "//div[@class='modal-content']//ul[contains(@class, 'nav')]/li[2]", "class").Contains("active"))
                    driver._click("xpath", "//div[@class='modal-content']//ul[contains(@class, 'nav')]/li[2]");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._waitForElement("xpath", "//mt-flyer-detail//img"), "Image not present in Detail Popup");
                Assert.IsTrue(driver._waitForElement("xpath", "//mt-flyer-detail//div[contains(@class, 'panel panel-box')]"), "Details not present in Detail Popup");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-header bgColor']/button"), "'X' button not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-footer']/button"), "'Close' button not present");
                Results.WriteStatus(test, "Pass", "Verified, Detail Popup.");
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "'Detail' Popup is still present.");
                Results.WriteStatus(test, "Pass", "Verified, Detail Popup is closed.");
            }
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Download Functionality in View Ad Popup
        ///</summary>
        ///<param name="buttonName">Download button to be clicked</param>
        ///<returns></returns>
        public AdDetails VerifyDownloadFunctionalityInViewAdPopup(string buttonName = "Download Full Ad")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "'View Ad' Popup not present.");
            if (!driver._getAttributeValue("xpath", "//div[@class='modal-content']//ul[contains(@class, 'nav')]/li[1]", "class").Contains("active"))
                driver._click("xpath", "//div[@class='modal-content']//ul[contains(@class, 'nav')]/li[1]");
            Thread.Sleep(1000);

            if (buttonName.ToLower().Contains("current"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//mt-flyer//button[contains(text(), 'Download Current Page')]"), "'" + buttonName + "' button not present");
                driver._click("xpath", "//mt-flyer//button[contains(text(), 'Download Current Page')]");
            }
            else if (buttonName.ToLower().Contains("selected"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//mt-flyer//div[contains(@class, 'active')]/ul/li//span"), "Checkbox on Pages not present.");
                IList<IWebElement> checkboxColl = driver._findElements("xpath", "//mt-flyer//div[contains(@class, 'active')]/ul/li//span");
                Random rand = new Random();
                int x = rand.Next(0, checkboxColl.Count);
                checkboxColl[x].Click();
                Assert.IsTrue(driver._waitForElement("xpath", "//mt-flyer//button[contains(text(), 'Download Selected Page')]"), "'" + buttonName + "' button not present");
                driver._click("xpath", "//mt-flyer//button[contains(text(), 'Download Selected Page')]");
            }
            else if (buttonName.ToLower().Contains("full"))
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//mt-flyer//button[contains(text(), 'Download Full Ad')]"), "'" + buttonName + "' button not present");
                driver._click("xpath", "//mt-flyer//button[contains(text(), 'Download Full Ad')]");
            }
            else
                Results.WriteStatus(test, "Fail", "Button not recognized.");

            Thread.Sleep(10000);
            Results.WriteStatus(test, "Pass", "Clicked, '" + buttonName + "' button in View Ad Popup.");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Pagination Functionality
        ///</summary>
        ///<param name="page">Page no. to be selected</param>
        ///<returns></returns>
        public AdDetails VerifyPaginationFunctionality(string page = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//ul[contains(@class, 'pagination')]/li/a"), "Pagination not present.");
            IList<IWebElement> paginationColl = driver._findElements("xpath", "//ul[contains(@class, 'pagination')]/li/a");

            if (page.ToLower().Contains("first"))
            {
                driver._scrollintoViewElement("xpath", "//ul[contains(@class, 'pagination')]/li/a[text()='First']");
                driver._click("xpath", "//ul[contains(@class, 'pagination')]/li/a[text()='First']");
                Thread.Sleep(1000);
                paginationColl = driver._findElements("xpath", "//ul[contains(@class, 'pagination')]/li");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a[text()='1']"), "First Page is not selected.");
                Assert.IsTrue(paginationColl[0].GetAttribute("class").Contains("disabled"), "First Page Button is not disabled.");
                Assert.IsTrue(paginationColl[1].GetAttribute("class").Contains("disabled"), "Previous Page Button is not disabled.");
            }
            else if (page.ToLower().Contains("last"))
            {
                driver._scrollintoViewElement("xpath", "//ul[contains(@class, 'pagination')]/li/a[text()='Last']");
                driver._click("xpath", "//ul[contains(@class, 'pagination')]/li/a[text()='Last']");
                Thread.Sleep(1000);
                paginationColl = driver._findElements("xpath", "//ul[contains(@class, 'pagination')]/li");
                Assert.IsTrue(paginationColl[paginationColl.Count - 3].GetAttribute("class").Contains("active"), "Last Page is not selected.");
                Assert.IsTrue(paginationColl[paginationColl.Count - 2].GetAttribute("class").Contains("disabled"), "Next Page Button is not disabled.");
                Assert.IsTrue(paginationColl[paginationColl.Count - 1].GetAttribute("class").Contains("disabled"), "Last Page Button is not disabled.");
            }
            else if (page.ToLower().Contains("prev"))
            {
                driver._scrollintoViewElement("xpath", "//ul[contains(@class, 'pagination')]/li/a[text()='First']");
                string currentPage = driver._getText("xpath", "//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a");
                int atPage = 0;
                Assert.IsTrue(int.TryParse(currentPage, out atPage), "Couldn't convert '" + currentPage + "' to int.");
                driver._click("xpath", "//ul[contains(@class, 'pagination')]/li[contains(@class, 'prev')]/a");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a[text()='" + (atPage - 1).ToString() + "']"), "Previous Page was not selected.");
            }
            else if (page.ToLower().Contains("next"))
            {
                driver._scrollintoViewElement("xpath", "//ul[contains(@class, 'pagination')]/li/a[text()='First']");
                string currentPage = driver._getText("xpath", "//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a");
                int atPage = 0;
                Assert.IsTrue(int.TryParse(currentPage, out atPage), "Couldn't convert '" + currentPage + "' to int.");
                driver._click("xpath", "//ul[contains(@class, 'pagination')]/li[contains(@class, 'next')]/a");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a[text()='" + (atPage + 1).ToString() + "']"), "Next Page was not selected.");
            }
            else
            {
                bool avail = false;
                IList<IWebElement> paginationCollList = driver._findElements("xpath", "//ul[contains(@class, 'pagination')]/li");
                while (!avail && !paginationCollList[paginationCollList.Count - 2].GetAttribute("class").Contains("disabled"))
                {
                    for (int i = 2; i < paginationColl.Count - 2; i++)
                        if (paginationColl[i].Text.Equals(page))
                        {
                            avail = true;
                            paginationColl[i].Click();
                            break;
                        }
                    if (avail)
                        break;
                    paginationColl[paginationColl.Count - 2].Click();
                    Thread.Sleep(1000);
                    paginationColl = driver._findElements("xpath", "//ul[contains(@class, 'pagination')]/li/a");
                    paginationCollList = driver._findElements("xpath", "//ul[contains(@class, 'pagination')]/li");
                }
                Assert.IsTrue(avail, "'" + page + "' not found.");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[contains(@class, 'pagination')]/li[contains(@class, 'active')]/a[text()='" + page + "']"), "'Page' was not selected.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Pagination functionality by navigating to '" + page + "' page.");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Items Per Page Functionality
        ///</summary>
        ///<param name="noOfItems">No of Items per page to be selected</param>
        ///<returns></returns>
        public AdDetails VerifyItemsPerPageFunctionality(string noOfItems = "")
        {
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'page-size')]//button/a"), "Items per page buttons not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='imageView']/div"), "Records not present.");
            IList<IWebElement> recordsColl = driver._findElements("xpath", "//div[@id='imageView']/div");
            string selectedItemsNum = driver._getText("xpath", "//div[contains(@class, 'page-size')]//button[contains(@class, 'active')]/a");
            Assert.LessOrEqual(recordsColl.Count.ToString(), selectedItemsNum, "Selected Items per page do not match the displayed no. of items.");

            if (selectedItemsNum.Equals(noOfItems))
                Results.WriteStatus(test, "Pass", "'" + noOfItems + "' Items per page is already selected.");
            else
            {
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@class, 'page-size')]//button[@btn-radio='" + noOfItems + "']"), "'" + noOfItems + "' Items per page button not present.");
                driver._click("xpath", "//div[contains(@class, 'page-size')]//button[@btn-radio='" + noOfItems + "']");
                homePage.VerifyHomePage();
                Thread.Sleep(2000);
                Assert.IsTrue(driver._getAttributeValue("xpath", "//div[contains(@class, 'page-size')]//button[@btn-radio='" + noOfItems + "']", "class").Contains("active"), "'" + noOfItems + "' Items per page button was not selected.");
                recordsColl = driver._findElements("xpath", "//div[@id='imageView']/div");
                Assert.LessOrEqual(recordsColl.Count.ToString(), noOfItems, "Selected Items per page do not match the displayed no. of items.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Items per page functionality by selecting to '" + noOfItems + "' Items per page.");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Verify Save Search Popup
        ///</summary>
        ///<param name="popupVisible">Whether popup should be visible</param>
        ///<returns></returns>
        public AdDetails VerifySaveSearchPopup(bool popupVisible = true)
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Save Search Popup is not present.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//h4"), "Save Search popup header is not present.");
                Assert.AreEqual("Save Search", driver._getText("xpath", "//div[@class='modal-content']//h4"), "Save Search Popup header text does not match.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[contains(text(), 'Search Name')]"), "Search Name Label not present in Save Search Popup.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[@id='searchableDDR']/div"), "Search Name Field not present in Save Search Popup.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[contains(text(), 'Visibility')]"), "Visibility Label not present in Save Search Popup.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[contains(@ng-click, 'eachVisibilityRadioObj')][1]"), "Private Radio Button not present in Save Search Popup.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[contains(@ng-click, 'eachVisibilityRadioObj')][2]"), "Public Radio Button not present in Save Search Popup.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//li[contains(text(), 'Do you want to make this your default search?')]"), "'Do you want to make this your default search?' Label not present in Save Search Popup.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[contains(@class, 'checkbox')]"), "'Do you want to make this your default search?' Checkbox not present in Save Search Popup.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-body-filters']//div[contains(@ng-class, 'AdvanceOptions')]"), "Advanced Options Button not present in Save Search Popup.");

                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-footer']//button[text()='Save Promo Search']"), "Save Promo Search Button not present in Save Search Popup.");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-footer']//button[text()='Cancel']"), "Cancel button not present in Save Search Popup.");
            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "Save Search Popup is still present.");
                Results.WriteStatus(test, "Pass", "Verified, Save Search Popup is closed.");
            }

            Results.WriteStatus(test, "Pass", "Verified, Save Search Popup");
            return new AdDetails(driver, test);
        }

        ///<summary>
        ///Edit Save Search Popup
        ///</summary>
        ///<returns></returns>
        public string EditSaveSearchPopup(string searchName = "", bool makePrivate = true, bool makeDefault = false, bool allData = true, bool bindSearch = false)
        {
            if (searchName != "")
            {
                if (searchName.ToLower().Equals("random"))
                    searchName = "TestSearch" + driver._randomString(4, true);

                //driver._click("xpath", "//div[@class='modal-body-filters']//div[@id='searchableDDR']/div");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//div[@class='modal-body-filters']//div[@id='searchableDDR']//input")).SendKeys(searchName);
                //driver._type("xpath", "//div[@class='modal-body-filters']//div[@id='searchableDDR']/div", searchName);
            }

            if (!makePrivate)
                driver._click("xpath", "//div[@class='modal-body-filters']//div[contains(@ng-click, 'eachVisibilityRadioObj')][2]//li");

            if (makeDefault)
                driver._click("xpath", "//div[@class='modal-body-filters']//div[contains(@ng-click, 'DefaultSearch')]//div[contains(@class, 'checkbox')]");

            if (!allData || bindSearch)
            {
                driver._click("xpath", "//div[@class='modal-body-filters']//div[contains(@ng-class, 'AdvanceOptions')]");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-body-filters']//div[contains(@ng-class, 'AdvanceOptions') and contains(@class, 'expanded')]"), "Advance Options did not expand.");

                if (!allData)
                    driver._click("xpath", "//div[@class='modal-body-filters']//div[contains(@ng-click, 'eachDataAdvancedSearch')][2]//li");

                if (bindSearch)
                    driver._click("xpath", "//div[@class='modal-body-filters']//div[contains(@ng-click, 'selectBind')]//div[contains(@class, 'checkbox')]");
            }

            Results.WriteStatus(test, "Pass", "Edited, Save Search Popup.");
            return searchName;
        }

        ///<summary>
        ///Verify Send Email As Popup
        ///</summary>
        ///<returns></returns>
        public AdDetails VerifySendEmailAsPopup(bool popupVisible = true, bool selectReportForEmail = false)
        {
            if (popupVisible)
            {
                Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='modal-content']"), "Send Email As Popup not present");
                Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='modal-content']//h4"), "Send Email As Popup header not present.");
                Assert.AreEqual("Send Email As", driver._getText("xpath", "//div[@class='modal-content']//h4"), "Send Email As Popup header text does not match.");

                Assert.IsTrue(driver._waitForElement("xpath", "//ul[contains(@class, 'nav nav-pills')]/li[@heading='Report']"), "Report Tab not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[contains(@class, 'nav nav-pills')]/li[@heading='Formatting Options']"), "Formatting Options Tab not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[contains(@class, 'nav nav-pills')]/li[@heading='Recipients']"), "Recipients Tab not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//ul[contains(@class, 'nav nav-pills')]/li[@heading='Email Options']"), "Email Options Tab not present.");

                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'modal-footer')]//button[text()='Email As Attachment']"), "Email As Attachment button not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'modal-footer')]//button[text()='Email with Download Link']"), "Email with Download Link button not present.");
                Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'modal-footer')]//button[text()='Cancel']"), "Cancel button not present.");

                if (selectReportForEmail)
                {
                    Assert.IsTrue(driver._isElementPresent("xpath", "//div[@id='leftcheckbox']//li"), "'Reports List' not present.");
                    driver._click("xpath", "//div[@id='leftcheckbox']//li[1]");
                    Thread.Sleep(1000);
                    IWebElement checkbox = driver.FindElement(By.XPath("//div[@id='leftcheckbox']//li[1]//input"));
                    Assert.AreNotEqual(null, checkbox, "Checkbox for 1st report not present");
                    Assert.AreNotEqual(null, checkbox.GetAttribute("checked"), "Checkbox for 1st report not checked.");
                }

            }
            else
            {
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//div[@class='modal-content']"), "Send Email As Popup still present");
                Results.WriteStatus(test, "Pass", "Verified, Send Email As Popup is closed.");
            }



            Results.WriteStatus(test, "Pass", "Verified, Send Email As Popup");
            return new AdDetails(driver, test);
        }










        #endregion
    }
}











