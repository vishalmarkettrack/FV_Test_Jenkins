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
    public class TestSuite017_FlashReports_RetailerComparison :Base
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
        FlashReports_RetailerComparison flashReports_RetailerComparison;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite017_FlashReports_RetailerComparison).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite017_FlashReports_RetailerComparison).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
            savedSearches = new SavedSearches(driver, test);
            calendar = new Calendar(driver, test);
            manufacturerComparison = new ManufacturerComparison(driver, test);
            flashReports_PromotedProducts = new FlashReports_PromotedProducts(driver, test);
            flashReports_RetailerComparison = new FlashReports_RetailerComparison(driver, test);
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC001");
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
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("FlashReports");
                homePage.VerifyHomePage();
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("", "Target");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifySubMenuOptionForClientAsFRData(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Sub Menu option for client as FR Data.");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifySubMenuOptionForClientAsFRData(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Sub Menu option for client as FR Data.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifyCountryDropDownList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify Country drop down list.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                flashReports_RetailerComparison.VerifyMedia_CountryDDL("Country", "United States");
                flashReports_RetailerComparison.VerifyMedia_CountryDDL("Country", "Canada");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyMediaDropDownList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify Media drop down list.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                flashReports_RetailerComparison.VerifyMedia_CountryDDL();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMedia_CountryDDL("Media", "Web");
                flashReports_RetailerComparison.VerifyMedia_CountryDDL("Media", "Email");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyCalendarTextboxAndIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify Calendar Textbox & icon.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyCalendarTextboxAndIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_VerifyExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify Export icon & options.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();

                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyExportOptions("PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Retailer-Comparison", "*.png");

                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyExportOptions("JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Retailer-Comparison", "*.jpg");

                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyExportOptions("PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Retailer-Comparison", "*.pdf");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_13_VerifyMyRetailerSectionAndAdInfo(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009_13-Verify My Retailer Section and Ad Info.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                flashReports_RetailerComparison.VerifyMyRetailerSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC009_13");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyMyRetailerSectionDropDownList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify MY RETAILER section Drop Down List.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerDropdown();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyMyRetailerSectionViewPageFullSizeText(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify MY RETAILER section VIEW PAGE FULL SIZE text.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyViewPageFullSizeLink();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyMyRetailerSectionPagination(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify MY RETAILER section Pagination.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection();
                flashReports_RetailerComparison.VerifyPagination("Next");
                flashReports_RetailerComparison.VerifyPagination("Prev");
                flashReports_RetailerComparison.VerifyPagination("3");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC012");
                throw;
            }
            driver.Quit();
        }
        
        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_18_VerifyCompetitorSectionAndAdInfo(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014_18-Verify COMPETITOR Section and Ad Info.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                flashReports_RetailerComparison.VerifyMyRetailerSection(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC014_18");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifyCompetitorSectionDropDownList(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify Competitor section Drop Down List.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection(false);
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerDropdown(false);
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifyCompetitorSectionViewPageFullSizeText(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify Competitor section VIEW PAGE FULL SIZE text.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection(false);
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyViewPageFullSizeLink(false);
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_VerifyCompetitorSectionPagination(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017-Verify Competitor section Pagination.");
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
                flashReports_PromotedProducts.VerifyAndSelectSubMenuOptionsOfFlashReport("Retailer Comparison");
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyRetailerComparisonScreen();
                driver._waitForElementToBeHidden("xpath", "//div[contains(@class, 'PopupLoader')]", 120);
                flashReports_RetailerComparison.VerifyMyRetailerSection();
                flashReports_RetailerComparison.VerifyPagination("Next", false);
                flashReports_RetailerComparison.VerifyPagination("Prev", false);
                flashReports_RetailerComparison.VerifyPagination("3", false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite017_FlashReports_RetailerComparison_TC017");
                throw;
            }
            driver.Quit();
        }
        
        #endregion
    }
}
