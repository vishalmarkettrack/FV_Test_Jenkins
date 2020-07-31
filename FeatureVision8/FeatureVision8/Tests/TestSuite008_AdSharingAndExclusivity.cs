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
    public class TestSuite008_AdSharingAndExclusivity : Base
    {
        string clientName = "Wells Dairy";

        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        AdDetails adDetails;
        Search seachPage;
        SavedSearches savedSearches;
        Calendar calendar;
        AdSharingAndExclusivity adSharingAndExclusivity;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite008_AdSharingAndExclusivity).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite008_AdSharingAndExclusivity).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
            savedSearches = new SavedSearches(driver, test);
            calendar = new Calendar(driver, test);
            adSharingAndExclusivity = new AdSharingAndExclusivity(driver, test);
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyAdSharingAndExclusivityPage(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify Ad Sharing And Exclusivity page.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsBrand(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Exclusive and Shared Ad Blocks tab DDL selected as Brand.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Brand");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsBrandHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Exclusive and Shared Ad Blocks tab DDL selected as Brand Help Icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Brand");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPageHelpIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsBrandWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify Exclusive and Shared Ad Blocks tab DDL selected as Brand Window Icon and Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Brand");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                adSharingAndExclusivity.VerifySortByNumberOrPercentageOnExclusiveAndSharedAdBlocks();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                adSharingAndExclusivity.VerifySortByNumberOrPercentageOnExclusiveAndSharedAdBlocks(true);

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Total Ad Boxes");
                adSharingAndExclusivity.VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Shared Ad Boxes");
                adSharingAndExclusivity.VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab("Brand", "Shared");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Exclusive Ad Boxes");
                adSharingAndExclusivity.VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab("Brand", "Exclusive");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsBrandExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify Exclusive and Shared Ad Blocks tab DDL selected as Brand Export Icon And Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Brand");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsCategory(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify Exclusive and Shared Ad Blocks tab DDL selected as Category.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Category");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsCategoryHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify Exclusive and Shared Ad Blocks tab DDL selected as Category Help Icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Category");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPageHelpIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsCategoryWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009-Verify Exclusive and Shared Ad Blocks tab DDL selected as Category Window Icon and Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Category");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                adSharingAndExclusivity.VerifySortByNumberOrPercentageOnExclusiveAndSharedAdBlocks();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                adSharingAndExclusivity.VerifySortByNumberOrPercentageOnExclusiveAndSharedAdBlocks(true);

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Total Ad Boxes");
                adSharingAndExclusivity.VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab("Category");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Shared Ad Boxes");
                adSharingAndExclusivity.VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab("Category", "Shared");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Exclusive Ad Boxes");
                adSharingAndExclusivity.VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab("Category", "Exclusive");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC009");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsCategoryExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify Exclusive and Shared Ad Blocks tab DDL selected as Category Export Icon And Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Category");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsManufacturer(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Exclusive and Shared Ad Blocks tab DDL selected as Manufacturer.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Manufacturer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsManufacturerHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify Exclusive and Shared Ad Blocks tab DDL selected as Manufacturer Help Icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Manufacturer");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPageHelpIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC013_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsManufacturerWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC013-Verify Exclusive and Shared Ad Blocks tab DDL selected as Manufacturer Window Icon and Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Manufacturer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by #");
                adSharingAndExclusivity.VerifySortByNumberOrPercentageOnExclusiveAndSharedAdBlocks();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Show by %");
                adSharingAndExclusivity.VerifySortByNumberOrPercentageOnExclusiveAndSharedAdBlocks(true);

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Total Ad Boxes");
                adSharingAndExclusivity.VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab("Manufacturer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Shared Ad Boxes");
                adSharingAndExclusivity.VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab("Manufacturer", "Shared");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Exclusive Ad Boxes");
                adSharingAndExclusivity.VerifySortByOptionsOnExclusiveAndSharedAdBlocksTab("Manufacturer", "Exclusive");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC013");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_VerifyExclusiveAndSharedAdBlocksTabDDLSelectedAsManufacturerExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014-Verify Exclusive and Shared Ad Blocks tab DDL selected as Manufacturer Export Icon And Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifyExclusiveAndSharedAdBlocksTab("Manufacturer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Exclusive and Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC014");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifySharedAdBlocksTabDDLSelectedAsBrand(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify Shared Ad Blocks tab DDL selected as Brand.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Brand");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifySharedAdBlocksTabDDLSelectedAsBrandHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify Shared Ad Blocks tab DDL selected as Brand Help Icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Brand");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPageHelpIcon("Shared Ad Blocks");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_VerifySharedAdBlocksTabDDLSelectedAsBrandWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017-Verify Shared Ad Blocks tab DDL selected as Brand Window Icon and Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Brand");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Total Ad Boxes");
                adSharingAndExclusivity.VerifySortByTotalAdBlocksOnSharedAdBlocksTab();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Shared Ad Boxes");
                adSharingAndExclusivity.VerifySortBySharedAdBlocksOnSharedAdBlocksTab();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC017");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC018_VerifySharedAdBlocksTabDDLSelectedAsBrandExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC018-Verify Shared Ad Blocks tab DDL selected as Brand Export Icon and Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Brand");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Brand Sort by Total Ad Boxes", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC018");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC019_VerifySharedAdBlocksTabDDLSelectedAsCategory(String Bname)
        {
            TestFixtureSetUp(Bname, "TC019-Verify Shared Ad Blocks tab DDL selected as Category.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Category");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC019");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC020_VerifySharedAdBlocksTabDDLSelectedAsCategoryHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC020-Verify Shared Ad Blocks tab DDL selected as Category Help Icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Category");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPageHelpIcon("Shared Ad Blocks");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC020");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC021_VerifySharedAdBlocksTabDDLSelectedAsCategoryWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC021-Verify Shared Ad Blocks tab DDL selected as Category Window Icon and Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Category");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Total Ad Boxes");
                adSharingAndExclusivity.VerifySortByTotalAdBlocksOnSharedAdBlocksTab();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Shared Ad Boxes");
                adSharingAndExclusivity.VerifySortBySharedAdBlocksOnSharedAdBlocksTab();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC021");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC022_VerifySharedAdBlocksTabDDLSelectedAsCategoryExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC022-Verify Shared Ad Blocks tab DDL selected as Category Export Icon and Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Category");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Category Sort by Total Ad Boxes", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC022");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC023_VerifySharedAdBlocksTabDDLSelectedAsManufacturer(String Bname)
        {
            TestFixtureSetUp(Bname, "TC023-Verify Shared Ad Blocks tab DDL selected as Manufacturer.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Manufacturer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC023");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC024_VerifySharedAdBlocksTabDDLSelectedAsManufacturerHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC024-Verify Shared Ad Blocks tab DDL selected as Manufacturer Help Icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Manufacturer");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPageHelpIcon("Shared Ad Blocks");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC024");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC025_VerifySharedAdBlocksTabDDLSelectedAsManufacturerWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC025-Verify Shared Ad Blocks tab DDL selected as Manufacturer Window Icon and Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Manufacturer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Total Ad Boxes");
                adSharingAndExclusivity.VerifySortByTotalAdBlocksOnSharedAdBlocksTab();

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("More Options", "Sort by Shared Ad Boxes");
                adSharingAndExclusivity.VerifySortBySharedAdBlocksOnSharedAdBlocksTab();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC025");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC026_VerifySharedAdBlocksTabDDLSelectedAsManufacturerExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC026-Verify Shared Ad Blocks tab DDL selected as Manufacturer Export Icon and Options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Manufacturer");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.png");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.jpeg");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.pdf");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.xlsx");

                adSharingAndExclusivity.SelectOptionFromDropDownOnAdSharingAndExclusivityPage("Export", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Shared Ad Blocks by Manufacturer Sort by Total Ad Boxes", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC026");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC027_VerifyCarouselWhenUserClickExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC027-Verify Carousel when user click Export icon and options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Manufacturer");
                calendar.VerifyExportAsExcelOptionFromCarousel();
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_Detail_Report_Ad_Blocks_", "*.xlsx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC027");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC028_VerifyCarouselWhenUserClickViewAdOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC028-Verify Carousel when user click View Ad option.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Manufacturer");
                calendar.OpenViewAdOrDetailPopup(false);
                adDetails.VerifyViewAdPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC028");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC029_VerifyCarouselWhenUserClickViewAdOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC029-Verify Carousel when user click View Ad option.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Manufacturer");
                calendar.OpenViewAdOrDetailPopup(true);
                adDetails.VerifyDetailPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC029");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC030_VerifyCarouselNavigation(String Bname)
        {
            TestFixtureSetUp(Bname, "TC030-Verify Carousel Navigation.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Category & Brand Share");
                adSharingAndExclusivity.VerifyAdSharingAndExclusivityPage();
                adSharingAndExclusivity.VerifySharedAdBlocksTab("Manufacturer");
                adSharingAndExclusivity.VerifyNavigationOnSharedAdBlocksTab();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite008_AdSharingAndExclusivity_TC030");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
