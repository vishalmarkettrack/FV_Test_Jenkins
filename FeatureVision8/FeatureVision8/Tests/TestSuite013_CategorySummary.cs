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
    public class TestSuite013_CategorySummary : Base
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        AdDetails adDetails;
        Search seachPage;
        SavedSearches savedSearches;
        Calendar calendar;
        AdSharingAndExclusivity adSharingAndExclusivity;
        CategorySummary categorySummary;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite013_CategorySummary).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite013_CategorySummary).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
            savedSearches = new SavedSearches(driver, test);
            calendar = new Calendar(driver, test);
            adSharingAndExclusivity = new AdSharingAndExclusivity(driver, test);
            categorySummary = new CategorySummary(driver, test);
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyCategorySummaryPageWhenClientTypeAsRetailer(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify Category Summary page when client type as Retailer.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyCategorySummaryPageWhenClientTypeAsRetailer(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Category Summary page when client type as Retailer.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("JCPenney");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("JCPenney");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifyCategorySummaryPageWhenClientTypeAsRetailerAndDBTypeAsAustralia(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Category Summary page when client type as Retailer And DB Type As Australia.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Metcash - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Metcash - Australia");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifyCategorySummaryPageWhenClientTypeAsManufacturer(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify Category Summary page when client type as Manufacturer.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyCategorySummaryPageWhenClientTypeAsManufacturerAndDBTypeAsDurable(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify Category Summary page when client type as Manufacturer And DB Type As Durable.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Activision");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Activision");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyCategorySummaryPageWhenClientTypeAsManufacturerAndDBTypeAsAustralia(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify Category Summary page when client type as Manufacturer And DB Type As Australia.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Mattel - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Mattel - Australia");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_VerifyCategorySummaryPageWhenClientTypeAsDistributor(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify Category Summary page when client type as Distributor..");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_VerifyCategorySummaryPageWhenClientTypeAsDistributorAndDBTypeAsDurable(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009-Verify Category Summary page when client type as Distributor & DB type as Durable.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Meyer Corporation");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Meyer Corporation");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC009");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyCategorySummaryPageWhenClientTypeAsDistributorAndDBTypeAsAustralia(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify Category Summary page when client type as Distributor & DB type as Australia.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Logitech - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Logitech - Australia");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyTop10RetailerTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Top 10 Retailer tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage();
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Retailer");
                categorySummary.VerifyTop10Tab();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyTop10RetailerTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify Top 10 Retailer tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage();
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Retailer");
                categorySummary.VerifyTop10Tab();
                categorySummary.VerifyCategorySummaryPageHelpIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC013_VerifyTop10RetailerTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC013-Verify Top 10 Retailer tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage();
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Retailer");
                categorySummary.VerifyTop10Tab();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC013");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_VerifyTop10RetailerTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014-Verify Top 10 Retailer tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage();
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Retailer");
                categorySummary.VerifyTop10Tab();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Retailer-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Retailer-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Retailer-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Retailer-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Retailer-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC014");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifyTop10BrandTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify Top 10 Brand tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("JCPenney");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("JCPenney");
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Brand");
                categorySummary.VerifyTop10Tab("Brand");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifyTop10BrandTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify Top 10 Brand tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("JCPenney");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("JCPenney");
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Brand");
                categorySummary.VerifyTop10Tab("Brand");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Top 10 Brand");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_VerifyTop10BrandTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017-Verify Top 10 Brand tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("JCPenney");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("JCPenney");
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Brand");
                categorySummary.VerifyTop10Tab("Brand");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC017");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC018_VerifyTop10BrandTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC018-Verify Top 10 Brand tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("JCPenney");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("JCPenney");
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Brand");
                categorySummary.VerifyTop10Tab("Brand");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Brand-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Brand-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Brand-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Brand-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Brand-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC018");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC019_VerifyPromotedProductsTabWithSelectedRadioAsChannel(String Bname)
        {
            TestFixtureSetUp(Bname, "TC019-Verify Promoted Products tab with selected Radio as Channel.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products");
                categorySummary.VerifyPromotedProductsTabWithRadioButton("Promoted Products by Channel", "Channel");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC019");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC020_VerifyPromotedProductsTabWithSelectedRadioAsChannelHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC020-Verify Promoted Products tab with selected Radio as Channel Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products");
                categorySummary.VerifyPromotedProductsTabWithRadioButton("Promoted Products by Channel", "Channel");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Products");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC020");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC021_VerifyPromotedProductsTabWithSelectedRadioAsChannelWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC021-Verify Promoted Products tab with selected Radio as Channel Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products");
                categorySummary.VerifyPromotedProductsTabWithRadioButton("Promoted Products by Channel", "Channel");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC021");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC022_VerifyPromotedProductsTabWithSelectedRadioAsChannelExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC022-Verify Promoted Products tab with selected Radio as Channel Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products");
                categorySummary.VerifyPromotedProductsTabWithRadioButton("Promoted Products by Channel", "Channel");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Channel-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Channel-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Channel-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Channel-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Channel-", "*.zip");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC022");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC023_VerifyPromotedProductsTabWithSelectedRadioAsParentRetailer(String Bname)
        {
            TestFixtureSetUp(Bname, "TC023-Verify Promoted Products tab with selected Radio as Parent Retailer.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products");
                categorySummary.VerifyPromotedProductsTabWithRadioButton("Promoted Products by Parent Retailer", "Parent Retailer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC023");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC024_VerifyPromotedProductsTabWithSelectedRadioAsParentRetailerHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC024-Verify Promoted Products tab with selected Radio as Parent Retailer Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products");
                categorySummary.VerifyPromotedProductsTabWithRadioButton("Promoted Products by Parent Retailer", "Parent Retailer");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Products");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC024");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC025_VerifyPromotedProductsTabWithSelectedRadioAsParentRetailerWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC025-Verify Promoted Products tab with selected Radio as Parent Retailer Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products");
                categorySummary.VerifyPromotedProductsTabWithRadioButton("Promoted Products by Parent Retailer", "Parent Retailer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC025");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC026_VerifyPromotedProductsTabWithSelectedRadioAsParentRetailerExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC026-Verify Promoted Products tab with selected Radio as Parent Retailer Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products");
                categorySummary.VerifyPromotedProductsTabWithRadioButton("Promoted Products by Parent Retailer", "Parent Retailer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Parent Retailer-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Parent Retailer-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Parent Retailer-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Parent Retailer-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Parent Retailer-", "*.zip");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC026");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC027_VerifyPromotedProductsByRetailerTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC027-Verify Promoted Products by Retailer tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Retailer");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Retailer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC027");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC028_VerifyPromotedProductsByRetailerTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC028-Verify Promoted Products by Retailer tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Retailer");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Retailer");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Product by Retailer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC028");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC029_VerifyPromotedProductsByRetailerTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC029-Verify Promoted Products by Retailer tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Retailer");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Retailer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC029");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC030_VerifyPromotedProductsByRetailerTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC030-Verify Promoted Products by Retailer tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Retailer");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Retailer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC030");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC031_VerifyPromotedProductsByRetailerGroupTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC031-Verify Promoted Products By Retailer Group tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Meyer Corporation");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Meyer Corporation");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Retailer Group");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Retailer Group");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC031");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC032_VerifyPromotedProductsByRetailerGroupTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC032-Verify Promoted Products By Retailer Group tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Meyer Corporation");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Meyer Corporation");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Retailer Group");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Retailer Group");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Product By Retailer Group");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC032");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC033_VerifyPromotedProductsByRetailerGroupTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC033-Verify Promoted Products By Retailer Group tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Meyer Corporation");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Meyer Corporation");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Retailer Group");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Retailer Group");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC033");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC034_VerifyPromotedProductsByRetailerGroupTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC034-Verify Promoted Products By Retailer Group tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Meyer Corporation");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Meyer Corporation");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Retailer Group");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Retailer Group");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer Group-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer Group-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer Group-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer Group-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Retailer Group-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC034");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC035_VerifyPromotedProductsByChannelTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC035-Verify Promoted Products by Channel tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Logitech - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Logitech - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Channel");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Channel");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC035");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC036_VerifyPromotedProductsByChannelTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC036-Verify Promoted Products by Channel tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Logitech - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Logitech - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Channel");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Channel");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Product by Channel");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC036");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC037_VerifyPromotedProductsByChannelTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC037-Verify Promoted Products by Channel tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Logitech - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Logitech - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Channel");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Channel");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC037");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC038_VerifyPromotedProductsByChannelTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC038-Verify Promoted Products by Channel tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Logitech - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Logitech - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Channel");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Channel");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Channel-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Channel-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Channel-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Channel-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Channel-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC038");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC039_VerifyPromotedProductsByManufacturerTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC039-Verify Promoted Products By Manufacturer tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage();
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products by Manufacturer");
                categorySummary.VerifyPromotedProductsTab("Promoted Products by Manufacturer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC039");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC040_VerifyPromotedProductsByManufacturerTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC040-Verify Promoted Products By Manufacturer tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage();
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products by Manufacturer");
                categorySummary.VerifyPromotedProductsTab("Promoted Products by Manufacturer");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Products by Manufacturer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC040");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC041_VerifyPromotedProductsByManufacturerTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC041-Verify Promoted Products By Manufacturer tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage();
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products by Manufacturer");
                categorySummary.VerifyPromotedProductsTab("Promoted Products by Manufacturer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC041");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC042_VerifyPromotedProductsByManufacturerTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC042-Verify Promoted Products By Manufacturer tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage();
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products by Manufacturer");
                categorySummary.VerifyPromotedProductsTab("Promoted Products by Manufacturer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Manufacturer-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Manufacturer-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Manufacturer-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Manufacturer-", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Manufacturer-", "*.zip");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC042");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC043_VerifyPromotedProductsByBrandTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC043-Verify Promoted Products by Brand tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("JCPenney");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("JCPenney");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products by Brand");
                categorySummary.VerifyPromotedProductsTab("Promoted Products by Brand");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC043");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC044_VerifyPromotedProductsByBrandTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC044-Verify Promoted Products by Brand tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("JCPenney");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("JCPenney");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products by Brand");
                categorySummary.VerifyPromotedProductsTab("Promoted Products by Brand");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Products by Brand");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC044");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC045_VerifyPromotedProductsByBrandTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC045-Verify Promoted Products by Brand tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("JCPenney");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("JCPenney");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products by Brand");
                categorySummary.VerifyPromotedProductsTab("Promoted Products by Brand");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC045");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC046_VerifyPromotedProductsByBrandTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC046-Verify Promoted Products by Brand tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("JCPenney");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("JCPenney");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Products by Brand");
                categorySummary.VerifyPromotedProductsTab("Promoted Products by Brand");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Brand", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Brand", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Brand", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Brand", "*.zip");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Products by Brand", "*.zip");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC046");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC047_VerifyTop10ManufacturerTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC047-Verify Top 10 Manufacturer tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Manufacturer");
                categorySummary.VerifyTop10Tab("Manufacturer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC047");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC048_VerifyTop10ManufacturerTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC048-Verify Top 10 Manufacturer tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Manufacturer");
                categorySummary.VerifyTop10Tab("Manufacturer");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Top 10 Manufacturer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC048");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC049_VerifyTop10ManufacturerTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC049-Verify Top 10 Manufacturer tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Manufacturer");
                categorySummary.VerifyTop10Tab("Manufacturer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC049");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC050_VerifyTop10ManufacturerTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC050-Verify Top 10 Manufacturer tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Top 10 Manufacturer");
                categorySummary.VerifyTop10Tab("Manufacturer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Manufacturer-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Manufacturer-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Manufacturer-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Manufacturer-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Top 10 Manufacturer-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC050");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC051_VerifyPromotedProductByOriginTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC051-Verify Promoted Product by Origin tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Origin");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Origin");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC051");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC052_VerifyPromotedProductByOriginTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC052-Verify Promoted Product by Origin tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Origin");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Origin");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Product by Origin");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC052");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC053_VerifyPromotedProductByOriginTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC053-Verify Promoted Product by Origin tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Origin");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Origin");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC053");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC054_VerifyPromotedProductByOriginTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC054-Verify Promoted Product by Origin tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Origin");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Origin");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Origin-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Origin-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Origin-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Origin-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Origin-", "*.png");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC054");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC055_VerifyPromotedProductByVarietyTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC055-Verify Promoted Product by Variety tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Variety");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Variety");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC055");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC056_VerifyPromotedProductByVarietyTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC056-Verify Promoted Product by Variety tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Variety");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Variety");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Product by Variety");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC056");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC057_VerifyPromotedProductByVarietyTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC057-Verify Promoted Product by Variety tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Variety");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Variety");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC057");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC058_VerifyPromotedProductByVarietyTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC058-Verify Promoted Product by Variety tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("California Table Grape");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("California Table Grape");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Variety");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Variety");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Variety-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Variety-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Variety-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Variety-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Variety-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC058");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC059_VerifyPromotedProductByCategoryTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC059-Verify Promoted Product by Category tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Meyer Corporation");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Meyer Corporation");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Category");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Category");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC059");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC060_VerifyPromotedProductByCategoryTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC060-Verify Promoted Product by Category tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Meyer Corporation");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Meyer Corporation");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Category");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Category");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Product by Category");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC060");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC061_VerifyPromotedProductByCategoryTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC061-Verify Promoted Product by Category tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Meyer Corporation");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Meyer Corporation");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Category");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Category");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC061");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC062_VerifyPromotedProductByCategoryTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC062-Verify Promoted Product by Category tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Meyer Corporation");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Meyer Corporation");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Category");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Category");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Category-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Category-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Category-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Category-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Category-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC062");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC063_VerifyPromotedProductBySubcategoryTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC063-Verify Promoted Product by Subcategory tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Logitech - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Logitech - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Subcategory");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Subcategory");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC063");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC064_VerifyPromotedProductBySubcategoryTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC064-Verify Promoted Product by Subcategory tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Logitech - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Logitech - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Subcategory");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Subcategory");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Promoted Product by Subcategory");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC064");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC065_VerifyPromotedProductBySubcategoryTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC065-Verify Promoted Product by Subcategory tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Logitech - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Logitech - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Subcategory");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Subcategory");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC065");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC066_VerifyPromotedProductBySubcategoryTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC066-Verify Promoted Product by Subcategory tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Logitech - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Logitech - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Promoted Product by Subcategory");
                categorySummary.VerifyPromotedProductsTab("Promoted Product by Subcategory");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Subcategory-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Subcategory-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Subcategory-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Subcategory-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Promoted Product by Subcategory-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC066");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC067_VerifyAdTypeByRetailersTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC067-Verify Ad Type by Retailers tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Ad Type by Retailers");
                categorySummary.VerifyAdType_MediumByRetailersTab();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC067");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC068_VerifyAdTypeByRetailersTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC068-Verify Ad Type by Retailers tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Ad Type by Retailers");
                categorySummary.VerifyAdType_MediumByRetailersTab();
                categorySummary.VerifyCategorySummaryPageHelpIcon("Ad Type by Retailers");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC068");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC069_VerifyAdTypeByRetailersTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC069-Verify Ad Type by Retailers tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Ad Type by Retailers");
                categorySummary.VerifyAdType_MediumByRetailersTab();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC069");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC070_VerifyAdTypeByRetailersTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC070-Verify Ad Type by Retailers tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Procter & Gamble");
                categorySummary.SelectTabOnCategorySummaryPage("Ad Type by Retailers");
                categorySummary.VerifyAdType_MediumByRetailersTab();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ad Type by Retailers-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ad Type by Retailers-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ad Type by Retailers-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ad Type by Retailers-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ad Type by Retailers-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC070");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC071_VerifyMediumByRetailersTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC071-Verify Medium by Retailers tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Metcash - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Metcash - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Medium by Retailers");
                categorySummary.VerifyAdType_MediumByRetailersTab("Medium");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC071");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC072_VerifyMediumByRetailersTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC072-Verify Medium by Retailers tab Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Metcash - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Metcash - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Medium by Retailers");
                categorySummary.VerifyAdType_MediumByRetailersTab("Medium");
                categorySummary.VerifyCategorySummaryPageHelpIcon("Medium by Retailers");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC072");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC073_VerifyMediumByRetailersTabWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC073-Verify Medium by Retailers tab Window Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Metcash - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Metcash - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Medium by Retailers");
                categorySummary.VerifyAdType_MediumByRetailersTab("Medium");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                categorySummary.VerifyShowByNumberOrPercentageOnCategorySummaryScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC073");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC074_VerifyMediumByRetailersTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC074-Verify Medium by Retailers tab Export Icon And Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Metcash - Australia");

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                categorySummary.VerifyCategorySummaryPage("Metcash - Australia");
                categorySummary.SelectTabOnCategorySummaryPage("Medium by Retailers");
                categorySummary.VerifyAdType_MediumByRetailersTab("Medium");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Medium by Retailers-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Medium by Retailers-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Medium by Retailers-", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Medium by Retailers-", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Medium by Retailers-", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite013_CategorySummary_TC074");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
