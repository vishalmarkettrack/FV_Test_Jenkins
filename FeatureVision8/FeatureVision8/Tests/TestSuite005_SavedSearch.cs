using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FeatureVision8.FeatureVision8
{
    [TestFixture]
    [Parallelizable(ParallelScope.Fixtures)]
    public class TestSuite005_SavedSearch : Base
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        AdDetails adDetails;
        Search seachPage;
        SavedSearches savedSearches;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite005_SavedSearch).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite005_SavedSearch).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
            savedSearches = new SavedSearches(driver, test);
            return driver;
        }

        [TearDown]
        public void TestFixtureTearDown()
        {
            extent.Flush();
            driver.Quit();
        }

        #endregion

        #region Test Methods

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC001_VerifyHomeScreenAfterLoginIntoApplication(String Bname)
        {
            TestFixtureSetUp(Bname, "TC001-Verify Home screen after login into Application.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyLeftNavigationMenuList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify Left Navigation Menu list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifySavedSearchesAndSummariesTabWithoutSavedList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Saved Searches and Summaries tab without saved list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifySavedSearchesAndSummariesTabWithSavedList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Saved Searches and Summaries tab with saved list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifyCreatedByDDL(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify Created By DDL.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifyCreatedByDropdown("Me");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyShowDDL(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify Show DDL.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifyShowDropdown("Shared");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifySearchBox(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify Search Box.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                string savedSearchName = savedSearches.getSavedSearchNameFromList();
                savedSearches.VerifySearchBox(savedSearchName);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_9_19_VerifyRunOptionForAnyOfTheSavedSearchListRecord(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008_9_19-Verify Run option for any of the Saved search list record.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                string savedSearchName = savedSearches.getSavedSearchNameFromList();
                savedSearches.VerifySearchBox(savedSearchName);
                savedSearches.VerifySavedSearchOptions("Run");
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC008_9_19");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyViewCriteriaOptionForAnyOfTheSavedSearchListRecord(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify View Criteria option for any of the Saved search list record.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                string savedSearchName = savedSearches.getSavedSearchNameFromList();
                savedSearches.VerifySearchBox(savedSearchName);
                savedSearches.VerifySavedSearchOptions("View Criteria");
                savedSearches.VerifyViewCriteriaWindow();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyManageLabelsOptionForAnyOfTheSavedSearchListRecord(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Manage Labels option for any of the Saved search list record.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                string savedSearchName = savedSearches.getSavedSearchNameFromList();
                savedSearches.VerifySearchBox(savedSearchName);
                string searchName = savedSearches.VerifySavedSearchOptions("Manage Label");
                savedSearches.VerifyManageLabelOptionsPopup(searchName);
                string newLabel = savedSearches.ChangeLabelFromManageLabelOptions();
                homePage.VerifyAlertPopupMessageAndClickButton("Label(s) has been updated successfully.", "Okay, Got It");
                savedSearches.VerifyLabelOfSavedSearches(newLabel);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyCreateSubscriptionButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify CREATE SUBSCRIPTION button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifyCreateSubscriptionPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_VerifyDragAndDropFunctionalityForAnyColumn(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014-Verify Drag & Drop functionality for any column.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifyDragAndDropFunctionalityOnSavedSearchColumns("Type", "Label");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC014");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifySavedSearchCreatedWithDDRIsGettingDisplayedWhileCreatingSubscription(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify saved search created with DDR is getting displayed while creating subscription.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Dynamic Date Range");
                seachPage.VerifySearchPageAndSelectCategory("Dynamic Date Range", null, "Month - Last 12 Calendar", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Save Search", "Save Search");
                adDetails.VerifySaveSearchPopup();
                string searchName = adDetails.EditSaveSearchPopup("Random");
                adDetails.ClickButtonInPopup("Save Promo Search");
                homePage.VerifyAlertPopupMessageAndClickButton("Promo Search \"" + searchName + "\" saved successfully.", "Okay, Got It");
                adDetails.VerifySaveSearchPopup(false);
                homePage.VerifyHomePage();
                homePage.VerifyHomeScreenInDetail(searchName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifyCreateSubscriptionPopup();
                savedSearches.SearchAndSelectSavedSearchInCreateSubscriptionPopup(searchName);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifySavedSearchCreatedWithoutDDRIsGettingDisplayedWhileCreatingSubscription(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify saved search created without DDR is getting displayed while creating subscription.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.RemoveAppliedSearchCriteriaFromMadlibSearch("Date");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Save Search", "Save Search");
                adDetails.VerifySaveSearchPopup();
                string searchName = adDetails.EditSaveSearchPopup("Random");
                adDetails.ClickButtonInPopup("Save Promo Search");
                homePage.VerifyAlertPopupMessageAndClickButton("Promo Search \"" + searchName + "\" saved successfully.", "Okay, Got It");
                adDetails.VerifySaveSearchPopup(false);
                homePage.VerifyHomePage();
                homePage.VerifyHomeScreenInDetail(searchName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifyCreateSubscriptionPopup();
                savedSearches.SearchAndSelectSavedSearchInCreateSubscriptionPopup(searchName, false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_VerifyTheOptionsForSavedSearchListWithCreatedSubscription(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017-Verify the options for Saved Search List with Created Subscription.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifyCreatedByDropdown("Me");
                homePage.VerifyHomePage();
                savedSearches.VerifySavedSearchOptions("", true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC017");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC018_VerifyTheOptionsForSavedSearchesNotCreatedByYou(String Bname)
        {
            TestFixtureSetUp(Bname, "TC018-Verify the options for Saved Searches not Created by you.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifyCreatedByDropdown("Numerator");
                homePage.VerifyHomePage();
                savedSearches.VerifySavedSearchOptions();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC018");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC020_VerifyTheOptionsForSavedSearchesNotCreatedByYou(String Bname)
        {
            TestFixtureSetUp(Bname, "TC020-Verify the options for Saved Searches not Created by you.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                homePage.VerifyHomePage();
                savedSearches.VerifySearchBox("Test Summary");
                string searchName = savedSearches.VerifySavedSearchOptions("Run");
                savedSearches.VerifySummaryScreen(searchName);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC020");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC021_VerifyWhenUserClickOnDeleteHavingNotSavedSubscription(String Bname)
        {
            TestFixtureSetUp(Bname, "TC021-Verify when user click on Delete having not saved subscription.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "February - 2020", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Save Search", "Save Search");
                adDetails.VerifySaveSearchPopup();
                string searchName = adDetails.EditSaveSearchPopup("Random");
                adDetails.ClickButtonInPopup("Save Promo Search");
                homePage.VerifyAlertPopupMessageAndClickButton("Promo Search \"" + searchName + "\" saved successfully.", "Okay, ");
                adDetails.VerifySaveSearchPopup(false);
                homePage.VerifyHomePage();
                homePage.VerifyHomeScreenInDetail(searchName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySearchBox(searchName);
                savedSearches.VerifySavedSearchOptions("Delete", true);
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to permanently delete this Promo Search?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("\"" + searchName + "\" Promo search deleted successfully. ", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC021");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC022_25_VerifyWhenUserClickOnDeleteHavingSavedSubscription(String Bname)
        {
            TestFixtureSetUp(Bname, "TC022_25-Verify when user click on Delete having saved subscription.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "February - 2020", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Save Search", "Save Search");
                adDetails.VerifySaveSearchPopup();
                string searchName = adDetails.EditSaveSearchPopup("Random");
                adDetails.ClickButtonInPopup("Save Promo Search");
                homePage.VerifyAlertPopupMessageAndClickButton("Promo Search \"" + searchName + "\" saved successfully.", "Okay, ");
                adDetails.VerifySaveSearchPopup(false);
                homePage.VerifyHomePage();
                homePage.VerifyHomeScreenInDetail(searchName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifyCreateSubscriptionPopup();
                savedSearches.SearchAndSelectSavedSearchInCreateSubscriptionPopup(searchName, true, true);
                adDetails.ClickButtonInPopup("Continue");
                string subsName = savedSearches.VerifyReportOptionsAndEditTabOfSubscription(searchName);
                savedSearches.VerifyAndEditScheduleAndFormatTabOfSubscription();
                savedSearches.ClickButtonOnSummaryScreen("Save & Close");
                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifySearchBox(searchName);
                savedSearches.VerifySubsColumnOnSavedSearchesScreen(searchName, "1");
                savedSearches.VerifySavedSearchOptions("Delete", true);
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to permanently delete this Promo Search?", "");
                homePage.VerifyAlertPopupMessageAndClickButton("If you delete this saved query it will be deleted from all subscription(s) that it is a part of, as follows:", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("\"" + searchName + "\" Promo search deleted successfully. ", "");
                homePage.VerifyAlertPopupMessageAndClickButton("have been deleted because they are no longer contain any saved queries.", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC022_25");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC023_VerifyWhenUserClickOnDeleteAndCancelButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC023-Verify when user click on Delete and Cancel button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "February - 2020", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Save Search", "Save Search");
                adDetails.VerifySaveSearchPopup();
                string searchName = adDetails.EditSaveSearchPopup("Random");
                adDetails.ClickButtonInPopup("Save Promo Search");
                homePage.VerifyAlertPopupMessageAndClickButton("Promo Search \"" + searchName + "\" saved successfully.", "Okay, ");
                adDetails.VerifySaveSearchPopup(false);
                homePage.VerifyHomePage();
                homePage.VerifyHomeScreenInDetail(searchName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifySearchBox(searchName);
                savedSearches.VerifySavedSearchOptions("Delete", true);
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to permanently delete this Promo Search?", "Cancel");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC023");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC024_VerifyWhenUserWantsToViewSubscription(String Bname)
        {
            TestFixtureSetUp(Bname, "TC024-Verify when user wants to View Subscription.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "February - 2020", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Save Search", "Save Search");
                adDetails.VerifySaveSearchPopup();
                string searchName = adDetails.EditSaveSearchPopup("Random");
                adDetails.ClickButtonInPopup("Save Promo Search");
                homePage.VerifyAlertPopupMessageAndClickButton("Promo Search \"" + searchName + "\" saved successfully.", "Okay, ");
                adDetails.VerifySaveSearchPopup(false);
                homePage.VerifyHomePage();
                homePage.VerifyHomeScreenInDetail(searchName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifyCreateSubscriptionPopup();
                savedSearches.SearchAndSelectSavedSearchInCreateSubscriptionPopup(searchName, true, true);
                adDetails.ClickButtonInPopup("Continue");
                string subsName = savedSearches.VerifyReportOptionsAndEditTabOfSubscription(searchName);
                savedSearches.VerifyAndEditScheduleAndFormatTabOfSubscription();
                savedSearches.ClickButtonOnSummaryScreen("Save & Close");
                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Saved Searches");
                savedSearches.VerifySavedSearchesScreen();
                savedSearches.VerifySearchBox(searchName);
                savedSearches.VerifySubsColumnOnSavedSearchesScreen(searchName, "1", true);
                savedSearches.VerifySubscriptionsForQueryPopup(searchName, new string[]{ subsName});
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite005_SavedSearch_TC024");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
