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
    public class TestSuite012_PricingSummary : Base
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        MyReports myReports;
        SavedSearches savedSearches;
        AdDetails adDetails;
        PricingAndPromotions pricingAndPromotions;
        Search searchPage;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite012_PricingSummary).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite012_PricingSummary).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            myReports = new MyReports(driver, test);
            savedSearches = new SavedSearches(driver, test);
            adDetails = new AdDetails(driver, test);
            pricingAndPromotions = new PricingAndPromotions(driver, test);
            searchPage = new Search(driver, test);

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
        public void TC001_VerifyPromoSearchScreenAfterLogin(String Bname)
        {
            TestFixtureSetUp(Bname, "TC001-Verify Promo Search screen after login.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite012_PricingSummary_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyWhenUserMadeAnySelectionsInMadlibSearch(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify when user made any selections in madlib search.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                homePage.verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen("Pricing & Promotions", "Pricing Summary");
                pricingAndPromotions.verifyPricingSummaryScreen();

                searchPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Dynamic Date Range");
                searchPage.VerifySearchPageAndSelectCategory("Dynamic Date Range", null, "", "Run Report");
                pricingAndPromotions.verifyPricingSummaryScreen();

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite012_PricingSummary_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyPricingSummaryPage(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Pricing Summary page.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                homePage.verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen("Pricing & Promotions", "Pricing Summary");
                pricingAndPromotions.verifyPricingSummaryScreen();

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite012_PricingSummary_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifyManufacturerSummaryTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Manufacturer Summary tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                homePage.verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen("Pricing & Promotions", "Pricing Summary");
                pricingAndPromotions.verifyPricingSummaryScreen();

                pricingAndPromotions.verifyTooltipTextOfHelpIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite012_PricingSummary_TC004");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
