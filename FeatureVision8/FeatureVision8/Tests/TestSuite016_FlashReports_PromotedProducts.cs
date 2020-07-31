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
    public class TestSuite016_FlashReports_PromotedProducts : Base
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        AdDetails adDetails;
        Search seachPage;
        SavedSearches savedSearches;
        Calendar calendar;
        ManufacturerComparison manufacturerComparison;
        FlashReports_PromotedProducts flashReports_PromotedProducts;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite016_FlashReports_PromotedProducts).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite016_FlashReports_PromotedProducts).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
            savedSearches = new SavedSearches(driver, test);
            calendar = new Calendar(driver, test);
            manufacturerComparison = new ManufacturerComparison(driver, test);
            flashReports_PromotedProducts = new FlashReports_PromotedProducts(driver, test);
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifySubMenuOptionForClientAsFRData(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify Sub Menu option for client as FR Data.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyPromotedProductsScreen(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Promoted Products screen.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen(); 
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifyExportOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Export Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifyPage_CountryDDL("Page", "Front");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                string[,] datagrid = flashReports_PromotedProducts.CaptureDataFromDetailDataTable();
                flashReports_PromotedProducts.VerifyExportOptions("Excel");
                string fileName = homePage.VerifyFileDownloadedOrNotOnScreen("FlashReport_PromotedProducts_", "*.xlsx");
                flashReports_PromotedProducts.VerifyDataFromTabularGridInExportedExcelFile(fileName, datagrid);

                flashReports_PromotedProducts.VerifyExportOptions("PPT");
                Thread.Sleep(60000);
                homePage.VerifyFileDownloadedOrNotOnScreen("FlashReport_PromotedProducts_", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifyPageDropDownList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify Page drop down list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifyPage_CountryDDL();
                flashReports_PromotedProducts.VerifyPage_CountryDDL("Page", "Front Page");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyCountryDropDownList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify Country drop down list.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifyPage_CountryDDL("Country", "Canada");
                flashReports_PromotedProducts.VerifyPage_CountryDDL("Country", "United States");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyCalendarTextboxAndIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify Calendar Textbox & Icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifyCalendarTextboxAndIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_VerifySectionWithTabs(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify section with tabs.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen();
                flashReports_PromotedProducts.VerifyTabsOnPromotedProductsPage();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifyTabsOnPromotedProductsPage("Promoted Product Images");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifyTabsOnPromotedProductsPage("Page Images");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_VerifyDetailDataTabSortOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009-Verify Detail Data tab Sort option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifySortingFunctionality("Ascending");
                flashReports_PromotedProducts.VerifySortingFunctionality("Descending");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC009");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyDetailDataTabImageIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify Detail Data tab Image icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifyProductImageOnDetailDataTab();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyPageNavigationOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Page Navigation options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifyPaginationFunctionality("Next");
                flashReports_PromotedProducts.VerifyPaginationFunctionality("Last");
                flashReports_PromotedProducts.VerifyPaginationFunctionality("Previous");
                flashReports_PromotedProducts.VerifyPaginationFunctionality("First");
                flashReports_PromotedProducts.VerifyPaginationFunctionality("3");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyShowPerPageOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify Show per Page options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyHomePage();
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Promoted Products");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyPromotedProductsScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_PromotedProducts.VerifyItemsPerPageFunctionality("20");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite016_FlashReports_PromotedProducts_TC012");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
