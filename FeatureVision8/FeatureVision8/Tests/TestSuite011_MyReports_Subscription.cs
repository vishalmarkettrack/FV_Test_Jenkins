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
    public class TestSuite011_MyReports_Subscription : Base
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        MyReports myReports;
        SavedSearches savedSearches;
        AdDetails adDetails;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite011_MyReports_Subscription).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite011_MyReports_Subscription).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            myReports = new MyReports(driver, test);
            savedSearches = new SavedSearches(driver, test);
            adDetails = new AdDetails(driver, test);

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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC001");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifySubscriptionsTabWithoutSavedList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Subscriptions tab without saved list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(true);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifySubscriptionsTabWithSavedList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Subscriptions tab with saved list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);
                myReports.verify_OR_Select_OptionsOfSubscriptionName();

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifySearchBox(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify Search box.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);
                myReports.verify_OR_Select_OptionsOfSubscriptionName();

                string savedResultName = myReports.getSavedResultNameFromList("Subscriptions");
                myReports.insertValueInSearchBoxAndVerifyWithSavedNameList(savedResultName, "Subscriptions");

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyWhenOptionsForAnySavedSubscriptionRecord(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify when options for any saved Subscription record.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);
                myReports.verify_OR_Select_OptionsOfSubscriptionName(true);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_TC013_Verify_VIEW_EDIT_OptionForScriptionRecord(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007_TC013-Verify VIEW / EDIT option for scription record.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);
                string subName = myReports.verify_OR_Select_OptionsOfSubscriptionName(true, "View / Edit");
                myReports.VerifyViewEditSubscriptionScreen(subName);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC007_TC0013");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_Verify_SEND_ME_NOW_ButtonFromSubscriptionPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify SEND ME NOW button from Subscription popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);

                string subName = myReports.verify_OR_Select_OptionsOfSubscriptionName(true, "Send Me Now");
                homePage.verifySuccessAlertPopupWindowWithMessage("Subscription has been sent.");

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_Verify_SEND_TO_ALL_ButtonFromSubscriptionPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009-Verify SEND TO ALL button from Subscription popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);

                string subName = myReports.verify_OR_Select_OptionsOfSubscriptionName(true, "Send To All");
                homePage.verifySuccessAlertPopupWindowWithMessage("Subscription has been sent to all recipients.");

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC009");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_Verify_CREATE_SUBSCRIPTION_Button(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify CREATE SUBSCRIPTION button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);

                savedSearches.VerifyCreateSubscriptionPopup();
                string queryName = myReports.selectSavedSearchQueryFromSubscriptionPopup();
                adDetails.ClickButtonInPopup("Continue");
                string subsName = savedSearches.VerifyReportOptionsAndEditTabOfSubscription(queryName);

                savedSearches.VerifyAndEditScheduleAndFormatTabOfSubscription();
                savedSearches.ClickButtonOnSummaryScreen("Save & Close");

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifySortingFunctionality(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Sorting functionality.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);

                myReports.clickColumnHeaderForSortGridAndVerifyColumnHeaderIcon("Subscription Name", true);
                myReports.clickColumnHeaderForSortGridAndVerifyColumnHeaderIcon("Subscription Name", false);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyDragAndDropFunctionalityForAnyColumn(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify Drag & Drop functionality for any column.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);

                savedSearches.VerifyDragAndDropFunctionalityOnSavedSearchColumns("Type", "Last Run");

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_TC015_VerifyWhenUserClickOnDeleteAndCancelButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014_TC015-Verify when user click on Delete and Cancel button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("Subscriptions");
                myReports.verifySubscriptionSectionInDetail(false);

                string subName = myReports.verify_OR_Select_OptionsOfSubscriptionName(true, "Delete");
                homePage.verifyAlertPopupWindowWithMessage("Are You sure you want to permanently delete the following subscription:(" + subName + ")?", "Cancel", "No Message");
                myReports.verifySavedResultNamePresentOrNotOnList("Subscriptions", subName, true);

                string newSubName = myReports.verify_OR_Select_OptionsOfSubscriptionName(true, "Delete");
                homePage.verifyAlertPopupWindowWithMessage("Are You sure you want to permanently delete the following subscription:(" + newSubName + ")?", "Ok", "Subscription Deleted Successfully.");
                myReports.verifySavedResultNamePresentOrNotOnList("Subscriptions", newSubName, false);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite011_MyReports_Subscription_TC014_TC015");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}


