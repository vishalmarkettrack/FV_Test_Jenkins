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
    public class TestSuite006_RetailerActivity : Base
    {
        string clientName = "Wells Dairy";

        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        AdDetails adDetails;
        Search seachPage;
        SavedSearches savedSearches;
        RetailerActivity retailerActivity;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite006_RetailerActivity).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite006_RetailerActivity).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
            savedSearches = new SavedSearches(driver, test);
            retailerActivity = new RetailerActivity(driver, test);
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyRetailerActivityScreenForUSClient(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify Retailer Activity screen for US client.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyRetailerActivityScreenForCanadaClient(String Bname)

        {
            TestFixtureSetUp(Bname, "TC003-Verify Retailer Activity screen for Canada client.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Cascades Canada");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen(true, "Cascades Canada");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifyRetailerActivityScreenForAustraliaClient(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Retailer Activity screen for Australia client.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Metcash - Australia");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen(true, "Metcash - Australia");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_8_VerifyAdsByChannelTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005_8-Verify Ads by Channel tab Export icon and options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection();

                retailerActivity.VerifyExportMenuAndSelectOption("Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Channel", "*.png");

                retailerActivity.VerifyExportMenuAndSelectOption("Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Channel", "*.jpeg");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Channel", "*.pdf");

                retailerActivity.VerifyExportMenuAndSelectOption("Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Channel", "*.xlsx");

                retailerActivity.VerifyExportMenuAndSelectOption("Download Powerpoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Channel", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC005_8");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyAdsByChannelTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify Ads by Channel tab Help icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");

                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection();
                retailerActivity.VerifyHelpIconOnAdsByChannelOrParentRetailerSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyAdsByChannelTabWindowIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify Ads by Channel tab Window icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection();

                retailerActivity.VerifyWindowIconOnAdsByChannelOrParentRetailerSection(false, "Show by %");
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection();

                retailerActivity.VerifyWindowIconOnAdsByChannelOrParentRetailerSection(false, "Show by #");
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection();

                retailerActivity.VerifyWhenColumnIsDrilledDownOn();
                retailerActivity.VerifyWindowIconOnAdsByChannelOrParentRetailerSection(true, "Reset");
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_12_VerifyAdsByChannelTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009_12-Verify Ads by Channel tab Export icon and options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection(false);

                retailerActivity.VerifyExportMenuAndSelectOption("Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Parent Retailer", "*.png");

                retailerActivity.VerifyExportMenuAndSelectOption("Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Parent Retailer", "*.jpeg");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Parent Retailer", "*.pdf");

                retailerActivity.VerifyExportMenuAndSelectOption("Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Parent Retailer", "*.xlsx");

                retailerActivity.VerifyExportMenuAndSelectOption("Download Powerpoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Parent Retailer", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC009_12");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyAdsByParentRetailerTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify Ads by Parent Retailer tab Help icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection(false);
                retailerActivity.VerifyHelpIconOnAdsByChannelOrParentRetailerSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyAdsByParentRetailerTabWindowIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Ads by Parent Retailer tab Window icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection(false);

                retailerActivity.VerifyWindowIconOnAdsByChannelOrParentRetailerSection(false, "Show by %");
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection(false);

                retailerActivity.VerifyWindowIconOnAdsByChannelOrParentRetailerSection(false, "Show by #");
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection(false);

                retailerActivity.VerifyWhenColumnIsDrilledDownOn();
                retailerActivity.VerifyWindowIconOnAdsByChannelOrParentRetailerSection(true, "Reset");
                retailerActivity.VerifyAdsByChannel_ParentRetailerSection(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC013_16_VerifyPagesByChannelTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC013_16-Verify Pages by Channel tab Export icon and options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.SelectTabOnRetailerActivity("Pages by Channel");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection();

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Channel", "*.png");

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Channel", "*.jpeg");

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Channel", "*.pdf");

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Channel", "*.xlsx");

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download Powerpoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Channel", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC013_16");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_VerifyPagesByChannelTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014-Verify Pages by Channel tab Help icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");

                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.SelectTabOnRetailerActivity("Pages by Channel");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection();
                retailerActivity.VerifyHelpIconOnPagesByChannelOrParentRetailerSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC014");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifyPagesByChannelTabWindowIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify Pages by Channel tab Window icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");

                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.SelectTabOnRetailerActivity("Pages by Channel");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection();

                retailerActivity.VerifyWindowIconOnPagesByChannelOrParentRetailerSection(false, "Show by %");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection();

                retailerActivity.VerifyWindowIconOnPagesByChannelOrParentRetailerSection(false, "Show by #");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection();

                retailerActivity.VerifyWhenColumnIsDrilledDownOnFromPagesByChannel_ParentRetailer();
                retailerActivity.VerifyWindowIconOnPagesByChannelOrParentRetailerSection(true, "Reset");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_20_VerifyPagesByChannelTabExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017_20-Verify Pages by Channel tab Export icon and options.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");
                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.SelectTabOnRetailerActivity("Pages by Channel");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection(false);

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Parent Retailer", "*.png");

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Parent Retailer", "*.jpeg");

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Parent Retailer", "*.pdf");

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Parent Retailer", "*.xlsx");

                retailerActivity.VerifyExportMenuAndSelectOptionOnPagesByChannel_ParentRetailer("Download Powerpoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Parent Retailer", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC017_20");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC018_VerifyPagesByParentRetailerTabHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC018-Verify Pages by Parent Retailer tab Help icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");

                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.SelectTabOnRetailerActivity("Pages by Channel");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection(false);
                retailerActivity.VerifyHelpIconOnPagesByChannelOrParentRetailerSection();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC018");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC019_VerifyPagesByParentRetailerTabWindowIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC019-Verify Pages by Parent Retailer tab Window icon.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");

                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.SelectTabOnRetailerActivity("Pages by Channel");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection(false);

                retailerActivity.VerifyWindowIconOnPagesByChannelOrParentRetailerSection(false, "Show by %");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection(false);

                retailerActivity.VerifyWindowIconOnPagesByChannelOrParentRetailerSection(false, "Show by #");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection(false);

                retailerActivity.VerifyWhenColumnIsDrilledDownOn();
                retailerActivity.VerifyWindowIconOnPagesByChannelOrParentRetailerSection(true, "Reset");
                retailerActivity.VerifyPagesByChannel_ParentRetailerSection(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC019");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC021_22_23_VerifyAdsByWeekTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC021_22_22-Verify Ads by Week Tab.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");

                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.SelectTabOnRetailerActivity("Ads by Week");

                retailerActivity.VerifyAds_PagesByWeekSection();
                string headerCol = retailerActivity.OpenAdDetailsPopup();
                retailerActivity.VerifyAdDetailsPopup(headerCol, true);
                retailerActivity.VerifyHelpIconOnAds_PagesByWeekSection();

                retailerActivity.VerifyExportMenuAndSelectOption("Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Week", "*.png");

                retailerActivity.VerifyExportMenuAndSelectOption("Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Week", "*.jpeg");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Week", "*.pdf");

                retailerActivity.VerifyExportMenuAndSelectOption("Download Excel");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Week", "*.xlsx");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Ads by Week", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC021_22_23");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC024_25_26_VerifyPagesByWeekTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC024_25_26-Verify Pages by Week Tab.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");

                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.SelectTabOnRetailerActivity("Ads by Week");

                retailerActivity.VerifyAds_PagesByWeekSection(false);
                string headerCol = retailerActivity.OpenAdDetailsPopup();
                retailerActivity.VerifyAdDetailsPopup(headerCol, true);
                retailerActivity.VerifyHelpIconOnAds_PagesByWeekSection();

                retailerActivity.VerifyExportMenuAndSelectOption("Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Week", "*.png");

                retailerActivity.VerifyExportMenuAndSelectOption("Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Week", "*.jpeg");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Week", "*.pdf");

                retailerActivity.VerifyExportMenuAndSelectOption("Download Excel");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Week", "*.xlsx");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Pages by Week", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC024_25_26");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC027_28_29_VerifyEventsByWeekTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC024_25_26-Verify Events by Week Tab.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");

                retailerActivity.VerifyRetailerActivityScreen();
                retailerActivity.SelectTabOnRetailerActivity("Events by Week");

                retailerActivity.VerifyEvents_ThemesByWeekSection();
                string headerCol = retailerActivity.OpenAdDetailsPopup();
                retailerActivity.VerifyAdDetailsPopup(headerCol, true);
                retailerActivity.VerifyHelpIconOnEvents_ThemesByWeekSection();

                retailerActivity.VerifyExportMenuAndSelectOption("Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Events by Week", "*.png");

                retailerActivity.VerifyExportMenuAndSelectOption("Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Events by Week", "*.jpeg");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Events by Week", "*.pdf");

                retailerActivity.VerifyExportMenuAndSelectOption("Download Excel");
                homePage.VerifyFileDownloadedOrNotOnScreen("Events by Week", "*.xlsx");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Events by Week", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC027_28_29");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC030_31_32_VerifyThemesByWeekTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC030_31_32-Verify Themes by Week Tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Cascades Canada");
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Retailer Activity");

                retailerActivity.VerifyRetailerActivityScreen(true, "Cascades Canada");
                retailerActivity.SelectTabOnRetailerActivity("Themes by Week");

                retailerActivity.VerifyEvents_ThemesByWeekSection(false);
                string headerCol = retailerActivity.OpenAdDetailsPopup();
                retailerActivity.VerifyAdDetailsPopup(headerCol, true);
                retailerActivity.VerifyHelpIconOnEvents_ThemesByWeekSection();

                retailerActivity.VerifyExportMenuAndSelectOption("Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Themes by Week", "*.png");

                retailerActivity.VerifyExportMenuAndSelectOption("Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen("Themes by Week", "*.jpeg");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen("Themes by Week", "*.pdf");

                retailerActivity.VerifyExportMenuAndSelectOption("Download Excel");
                homePage.VerifyFileDownloadedOrNotOnScreen("Themes by Week", "*.xlsx");

                retailerActivity.VerifyExportMenuAndSelectOption("Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen("Themes by Week", "*.pptx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite006_RetailerActivity_TC030_31_32");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
