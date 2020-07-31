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
    public class TestSuite014_ManufacturerComparison : Base 
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        AdDetails adDetails;
        Search seachPage;
        SavedSearches savedSearches;
        ManufacturerComparison manufacturerComparison;
        AdSharingAndExclusivity adSharingAndExclusivity;
        CategorySummary categorySummary;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite014_ManufacturerComparison).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite014_ManufacturerComparison).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
            savedSearches = new SavedSearches(driver, test);
            manufacturerComparison = new ManufacturerComparison(driver, test);
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC001");
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

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "February - 2020", "Run Report");
                homePage.VerifyHomePage();
                manufacturerComparison.VerifyManufacturerComparisonPage(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyManufacturerComparisonPage(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Manufacturer Comparison Page.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifyManufacturerComparisonPageHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Manufacturer Comparison Page Help Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                manufacturerComparison.VerifyManufacturerComparisonPageHelpIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifyManufacturerComparisonPageExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify Manufacturer Comparison Page Export icon and options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Manufacturer Comparison-", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Manufacturer Comparison-", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Manufacturer Comparison-", "*.pdf");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyAnyColumnDropDownListIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify any column Drop Down List icon and options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                manufacturerComparison.VerifyColumnDDLOnManufacturerComparisonGrid();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyWhenUserClickAnyColumnNumberOfProducts(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify when user click any column # Products.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                manufacturerComparison.VerifyDetailDataOnManufacturerComparisonPage();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_VerifyDetailDataSectionRadioOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify Detail Data section Radio options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                manufacturerComparison.VerifyDetailDataOnManufacturerComparisonPage();
                manufacturerComparison.VerifyRadioButtonsOnDetailDataSection("Page Images");
                manufacturerComparison.VerifyRadioButtonsOnDetailDataSection("Promoted Product Images");
                manufacturerComparison.VerifyRadioButtonsOnDetailDataSection("Detail Data");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_VerifyDetailDataSectionExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009-Verify Detail Data section Export icon and options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                manufacturerComparison.VerifyDetailDataOnManufacturerComparisonPage();
                string[,] dataGrid = manufacturerComparison.CaptureDataFromDetailDataTable();
                manufacturerComparison.VerifyExportAsExcelOptionFromDetailDataSection();
                string fileName = homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_Detail_Report_Promoted_Product", "*.xlsx");
                manufacturerComparison.VerifyDataFromTabularGridInExportedExcelFile(fileName, dataGrid);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC009");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyDetailDataSectionPageNavigationOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify Detail Data section Page Navigation options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                manufacturerComparison.VerifyDetailDataOnManufacturerComparisonPage();
                adDetails.VerifyPaginationFunctionality("Next");
                adDetails.VerifyPaginationFunctionality("Last");
                adDetails.VerifyPaginationFunctionality("Previous");
                adDetails.VerifyPaginationFunctionality("First");
                adDetails.VerifyPaginationFunctionality("3");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyDetailDataSectionShowPerPageOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Detail Data section Show per Page options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                manufacturerComparison.VerifyDetailDataOnManufacturerComparisonPage();
                manufacturerComparison.VerifyRadioButtonsOnDetailDataSection("Page Images");
                manufacturerComparison.VerifyItemsPerPageFunctionality("40");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyNormalFilterOptionInAvailableColumns(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify Normal Filter option in available columns.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                manufacturerComparison.VerifyDetailDataOnManufacturerComparisonPage();
                manufacturerComparison.VerifyFilterFunctionalityOnDetailDataTable();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC013_VerifySortOptionInAvailableColumns(String Bname)
        {
            TestFixtureSetUp(Bname, "TC013-Verify Sort option in available columns.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                manufacturerComparison.VerifyManufacturerComparisonPage();
                manufacturerComparison.VerifyDetailDataOnManufacturerComparisonPage();
                manufacturerComparison.VerifySortFunctionalityOnDetailDataSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite014_ManufacturerComparison_TC013");
                throw;
            }
            driver.Quit();
        }



        #endregion
    }
}
