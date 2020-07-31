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
    public class TestSuite019_CustomReports_ExecutiveDashboard : Base
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        MyReports myReports;
        ExecutiveDashboard executiveDashboard;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite019_CustomReports_ExecutiveDashboard).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite019_CustomReports_ExecutiveDashboard).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            executiveDashboard = new ExecutiveDashboard(driver, test);

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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite019_CustomReports_ExecutiveDashboard_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyExecutiveDarshboradScreen(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify Executive Dashborad Screen.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.SelectOptionFromLeftNavigationMenuList("Custom Reports");
                homePage.verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen("Custom Reports", "Executive Dashboard");

                executiveDashboard.verifyExecutiveDashboardScreen();
                executiveDashboard.clickButtonAndSelectOptionFromList("Total Alcohol Beverage Ad Blocks", "Export", "Download PNG");
                executiveDashboard.VerifyFileDownloadedOrNotOnScreen("Total Alcohol Beverage Ad Blocks", "*.png");

                executiveDashboard.clickButtonAndSelectOptionFromList("Total Alcohol Beverage Ad Blocks", "Export", "Download JPG");
                executiveDashboard.VerifyFileDownloadedOrNotOnScreen("Total Alcohol Beverage Ad Blocks", "*.jpeg");

                executiveDashboard.clickButtonAndSelectOptionFromList("Total Alcohol Beverage Ad Blocks", "Export", "Download PDF");
                executiveDashboard.VerifyFileDownloadedOrNotOnScreen("Total Alcohol Beverage Ad Blocks", "*.pdf");

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite019_CustomReports_ExecutiveDashboard_TC002");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}



