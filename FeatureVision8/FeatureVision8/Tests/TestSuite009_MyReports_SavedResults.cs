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
    public class TestSuite009_MyReports_SavedResults : Base
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        MyReports myReports;
        SavedSearches savedSearches;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite009_MyReports_SavedResults).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite009_MyReports_SavedResults).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            myReports = new MyReports(driver, test);

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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyMyReportsScreen(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify My Reports screen.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifySavedResultsTabWithoutSavedList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Saved Results tab without saved list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen(true);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifySavedResultsTabWithSavedList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Saved Results tab with saved list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC004");
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
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                myReports.VerifyCreatedByDropdown("Me");

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifySearchBox(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify Search box.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                string savedResultName = myReports.getSavedResultNameFromList();
                myReports.insertValueInSearchBoxAndVerifyWithSavedNameList(savedResultName);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyTooltipAndClickOnAnySavedResultsHavingNoRecords(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify tooltip and click on any Saved Results having no records.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                myReports.VerifySavedResultsOptions("Run");
                homePage.VerifyHomeScreenInDetail("", false);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_VerifyTooltipAndClickOnAnySavedResultsHavingRecords(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify tooltip and click on any Saved Results having records.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                myReports.VerifySavedResultsOptions("Run");
                homePage.VerifyHomeScreenInDetail("", false);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_VerifySortingFunctionality(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009-Verify Sorting functionality.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                myReports.VerifySavedResultsOptions("Run");
                homePage.VerifyHomeScreenInDetail("", false);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC009");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyDragAndDropFunctionalityForAnyColumn(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify Sorting functionality.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                myReports.VerifyDragAndDropFunctionalityOnSavedResultsColumns("Label", "Type");

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyPlayIconWhenSavedResultsCreatedByClient(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Play icon when Saved Results created by client.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                myReports.verifySavedResultsList_CreatedBy_SearchName_FromList(false);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyPlayIconWhenSavedResultsCreatedByMe(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify Play icon when Saved Results created by me.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                myReports.verifySavedResultsList_CreatedBy_SearchName_FromList(true);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC013_VerifyWhenUserClickOnRunHavingTypeAsAdIndex(String Bname)
        {
            TestFixtureSetUp(Bname, "TC013-Verify when user click on Run having Type as Ad Index.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                myReports.verifyAndRunSelectedTypeSavedResults("Ad Index", "Run");
                homePage.VerifyHomeScreenInDetail("", true, true);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC013");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_VerifyWhenUserClickOnRunHavingTypeAsProductDetail(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014-Verify when user click on Run having Type as Product Detail.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();
                myReports.verifyAndRunSelectedTypeSavedResults("Product Detail", "Run");
                homePage.VerifyHomeScreenInDetail("", true, false);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC014");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifyWhenUserClickOnDeleteAndCancelButtonFromPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify when user click on Delete and Cancel button from popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();

                string SearchName = myReports.verifySavedResultsList_CreatedBy_SearchName_FromList(true, "Delete");
                homePage.verifyAlertPopupWindowWithMessage("Are you sure you want to delete the Saved result '" + SearchName + "'?", "Ok", "Recordset deleted successfully.");
                myReports.verifySavedResultNamePresentOrNotOnList("My Reports", SearchName, false);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifyWhenUserClickOnDeleteOptionForSavedResultCreatedByMe(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify when user click on Delete option for saved result created by Me.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();

                string SearchName = myReports.verifySavedResultsList_CreatedBy_SearchName_FromList(true, "Delete");
                homePage.verifyAlertPopupWindowWithMessage("Are you sure you want to delete the Saved result '" + SearchName + "'?", "Cancel", "No Message");
                myReports.verifySavedResultNamePresentOrNotOnList("My Reports", SearchName, true);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_VerifyWhenUserClickOnManageLabel(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017-Verify when user click on Manage Label.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");
                myReports.verifySavedResultsScreen();

                string SearchName = myReports.verifySavedResultsList_CreatedBy_SearchName_FromList(false, "Manage Label");
                savedSearches.VerifyManageLabelOptionsPopup(SearchName);
                string newLabel = savedSearches.ChangeLabelFromManageLabelOptions();
                homePage.VerifyAlertPopupMessageAndClickButton("Label(s) has been updated successfully.", "Okay, Got It");
                savedSearches.VerifyLabelOfSavedSearches(newLabel);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite009_MyReports_SavedResults_TC017");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
