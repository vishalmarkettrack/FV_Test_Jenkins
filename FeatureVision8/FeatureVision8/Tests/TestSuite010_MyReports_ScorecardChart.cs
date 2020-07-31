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
    public class TestSuite010_MyReports_ScorecardChart : Base
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
            Results.WriteTestSuiteHeading(typeof(TestSuite010_MyReports_ScorecardChart).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite010_MyReports_ScorecardChart).Name);

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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite010_MyReports_ScorecardChart_TC001");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite010_MyReports_ScorecardChart_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyScoreCardTabWithoutSavedList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify ScoreCard tab without saved list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("My Reports");

                myReports.verifySavedResultsScreen(true, false);
                myReports.selectAndVerifyTabFromMyReportsSection("ScoreCard Chart");
                myReports.verifySubscriptionSectionInDetail(true);

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite010_MyReports_ScorecardChart_TC003");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
