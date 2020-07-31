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
    public class Search
    {
        #region Private Variables

        private IWebDriver searchPage;
        private ExtentTest test;

        #endregion

        #region Public Methods

        public Search(IWebDriver driver, ExtentTest testReturn)
        {
            // TODO: Complete member initialization
            this.searchPage = driver;
            test = testReturn;
        }

        public IWebDriver driver
        {
            get { return this.searchPage; }
            set { this.searchPage = value; }
        }

        ///<summary>
        ///Verify Date Page And Select Date Category
        ///</summary>
        ///<param name="pageName">Search page</param>
        ///<param name="categoryNameList">To verify available categaries on pageName</param>
        ///<param name="category">Category to drill down on</param>
        ///<param name="button">Button to be clicked</param>
        ///<returns></returns>
        public Search VerifySearchPageAndSelectCategory(string pageName, string[] categoryNameList = null, string category = "", string button = "")
        {
            Thread.Sleep(2000);
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid]"), "Madlib Search Parameters not present.");
            if (pageName.ToLower().Equals("product"))
            {
                driver._click("xpath", "//div[@class='prompt-summary']//div[@madlibid='1']");
                RemoveAppliedSearchCriteriaFromMadlibSearch("Product");
            }
            else if (pageName.ToLower().Contains("account") || pageName.ToLower().Contains("retailer"))
            {
                driver._click("xpath", "//div[@class='prompt-summary']//div[@madlibid='2']");
                RemoveAppliedSearchCriteriaFromMadlibSearch("Account");
            }
            else if (pageName.ToLower().Contains("market"))
            {
                driver._click("xpath", "//div[@class='prompt-summary']//div[@madlibid='3']");
                RemoveAppliedSearchCriteriaFromMadlibSearch("Market");
            }
            else if (pageName.ToLower().Contains("date"))
            {
                driver._click("xpath", "//div[@class='prompt-summary']//div[@madlibid='4']");
                RemoveAppliedSearchCriteriaFromMadlibSearch("Date");
            }

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='prompt-container']"), "Search Page not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//input[@aria-label='search-input']"), "'Search' textbox not present.");
            Assert.AreEqual("Search", driver._getAttributeValue("xpath", "//input[@aria-label='search-input']", "placeholder"), "'Search' Text not present in Search Textbox.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[contains(@id, 'queryctr')]//div[@class='prompt-breadcrumbs']//li/span"), "'Breadcrumb' not present.");
            IList<IWebElement> breadCrumbCollection = driver._findElements("xpath", "//div[contains(@id, 'queryctr')]//div[@class='prompt-breadcrumbs']//li/span");
            Assert.IsTrue(breadCrumbCollection[breadCrumbCollection.Count - 1].Text.ToLower().Contains(pageName.ToLower()), "Search Page is not of '" + pageName + "'.");

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-container']//div[@class='prompt-element-label']//li"), "Categories not present.");
            IList<IWebElement> categoryCollection = driver._findElements("xpath", "//div[@class='prompt-container']//div[@class='prompt-element-label']//li");

            if(categoryNameList != null)
            {
                foreach (string categoryName in categoryNameList)
                {
                    bool avail = false;
                    foreach (IWebElement catEle in categoryCollection)
                        if (catEle.Text.ToLower().Equals(categoryName.ToLower()))
                        {
                            avail = true;
                            break;
                        }
                    Assert.IsTrue(avail, "'" + categoryName + "' Category not found.");
                }
            }

            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-buttons']//button[@ng-click='runQuery();']"), "Run Report button not present.");
            Assert.AreEqual("Run Report", driver._getText("xpath", "//div[@class='report-buttons']//button[@ng-click='runQuery();']/span"), "Run Report button text does not match.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='report-buttons']//button[@ng-click='hideSearchAreaAndMadLibsFieldsValuesList(true, false);']"), "Cancel button not present.");
            Assert.AreEqual("Cancel", driver._getText("xpath", "//div[@class='report-buttons']//button[@ng-click='hideSearchAreaAndMadLibsFieldsValuesList(true, false);']/span"), "Cancel button text does not match.");

            if (category != "")
            {
                bool avail = false;
                foreach (IWebElement catEle in categoryCollection)
                    if (catEle.Text.ToLower().Equals(category.ToLower()))
                    {
                        avail = true;
                        catEle.Click();
                        break;
                    }

                Assert.IsTrue(avail, "'" + category + "' Category not found.");
                Thread.Sleep(1000);
                Results.WriteStatus(test, "Pass", "Selected, '" + category + "' Category on Date Page");
            }
            else
            {
                Random rand = new Random();
                int x = rand.Next(0, categoryCollection.Count);
                categoryCollection[x].Click();
            }

            if (button.ToLower().Contains("run"))
                driver._click("xpath", "//div[@class='report-buttons']//button[@ng-click='runQuery();']");
            else if (button.ToLower().Equals("cancel"))
                driver._click("xpath", "//div[@class='report-buttons']//button[@ng-click='hideSearchAreaAndMadLibsFieldsValuesList(true);']");

            Results.WriteStatus(test, "Pass", "Verified, '" + pageName + "' Search Page.");
            return new Search(driver, test);
        }

        ///<summary>
        ///Remove Applied Search Criteria From Madlib Search
        ///</summary>
        ///<returns></returns>
        public Search RemoveAppliedSearchCriteriaFromMadlibSearch(string criteria)
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='prompt-summary']//div[@madlibid]"), "Madlib Search Parameters not present.");
            string madlibId = "";

            if (criteria.ToLower().Contains("product"))
                madlibId = "1";
            else if (criteria.ToLower().Contains("account") || criteria.ToLower().Contains("retailer"))
                madlibId = "2";
            else if (criteria.ToLower().Contains("market"))
                madlibId = "3";
            else if (criteria.ToLower().Contains("date"))
                madlibId = "4";

            if(driver._waitForElement("xpath", "//div[@class='prompt-summary']//div[@madlibid='" + madlibId + "']//i"))
            {
                driver._click("xpath", "//div[@class='prompt-summary']//div[@madlibid='" + madlibId + "']//i");
                Thread.Sleep(1000);
                Assert.IsTrue(driver._getText("xpath", "//div[@class='prompt-summary']//div[@madlibid='" + madlibId + "']").Contains("Any"), "Applied Search Criteria '" + criteria + "' did not get removed");
            }
            else
            {
                Assert.IsTrue(driver._getText("xpath", "//div[@class='prompt-summary']//div[@madlibid='" + madlibId + "']").Contains("Any"), "Applied Search Criteria '" + criteria + "' is set but cannot be removed.");
                Results.WriteStatus(test, "Pass", "No Search Criteria was applied to '" + criteria + "'");
                return new Search(driver, test);
            }

            Results.WriteStatus(test, "Pass", "Removed, Applied Search Criteria '" + criteria + "' From Madlib Search.");
            return new Search(driver, test);
        }

        ///<summary>
        ///Verify And Edit More Options In Search Criteria
        ///</summary>
        ///<returns></returns>
        public Search VerifyAndEditMoreOptionsInSearchCriteria(string optionName = "")
        {
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]"), "More Options not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/i"), "More Options Expand/Contract arrow not present.");
            Assert.IsTrue(driver._isElementPresent("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/span"), "No. of More Options selected not present.");
            string previousNumber = driver._getText("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/span");
            int prevNum = 0;
            previousNumber = previousNumber.TrimEnd(' ');
            previousNumber = previousNumber.TrimEnd(')');
            previousNumber = previousNumber.TrimStart(' ');
            previousNumber = previousNumber.TrimStart('(');
            Assert.IsTrue(int.TryParse(previousNumber, out prevNum), "Couldn't convert '" + previousNumber + "' to int");

            if (driver._getAttributeValue("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/i", "class").Contains("down"))
                driver._click("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]");

            Assert.IsTrue(driver._waitForElement("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/i[contains(@class, 'up')]"), "More Options not expanded.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'filter-block')]//div[@class='filter-label']/div"), "Option Names not present.");
            Assert.IsTrue(driver._waitForElement("xpath", "//div[contains(@class, 'filter-block')]//div[@class='prompt-answer-container']"), "Option Containers not present");

            IList<IWebElement> optionNameColl = driver._findElements("xpath", "//div[contains(@class, 'filter-block')]//div[@class='filter-label']/div");
            IList<IWebElement> optionFieldColl = driver._findElements("xpath", "//div[contains(@class, 'filter-block')]//div[@class='prompt-answer-container']");

            if(optionName == "" || optionName.ToLower().Equals("random"))
            {
                Random rand = new Random();
                int x = rand.Next(0, optionNameColl.Count);
                optionName = optionNameColl[x].Text;
            }

            bool avail = false;
            for(int i = 0;  i < optionNameColl.Count; i++)
                if (optionNameColl[i].Text.ToLower().Equals(optionName.ToLower() + ":"))
                {
                    avail = true;
                    IList<IWebElement> displayedTextColl = optionFieldColl[i]._findElementsWithinElement("xpath", ".//span");
                    if(!(displayedTextColl[0].Text.ToLower().Contains("any") 
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

                    IList<IWebElement> answerColl = optionFieldColl[i]._findElementsWithinElement("xpath", ".//div[contains(@class, 'answer selectable')]");
                    answerColl[0].Click();
                    Thread.Sleep(2000);
                    if (optionName.Contains(":"))
                        optionName = optionName.Substring(0, optionName.Length - 2);
                    VerifySearchPageAndSelectCategory(optionName, null, "", "Run Report");
                    Thread.Sleep(2000);
                    string newNumber = driver._getText("xpath", "//div[@class='header']/div[contains(text(), 'More Options')]/span");
                    int newNum = 0;
                    newNumber = newNumber.TrimEnd(' ');
                    newNumber = newNumber.TrimEnd(')');
                    newNumber = newNumber.TrimStart(' ');
                    newNumber = newNumber.TrimStart('(');
                    Assert.IsTrue(int.TryParse(newNumber, out newNum), "Couldn't convert '" + newNumber + "' to int");
                    Assert.AreNotEqual(newNum, prevNum, "Selected More Options number did not change");
                    break;
                }
            Assert.IsTrue(avail, "'" + optionName + "' not found.");

            Results.WriteStatus(test, "Pass", "Verified, And Edited '" + optionName + "' In Search Criteria.");
            return new Search(driver, test);
        }




        #endregion
    }
}
