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
    public class TestSuite015_PricingHistograms : Base
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
        PricingHistograms pricingHistograms;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite015_PricingHistograms).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite015_PricingHistograms).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
            savedSearches = new SavedSearches(driver, test);
            calendar = new Calendar(driver, test);
            manufacturerComparison = new ManufacturerComparison(driver, test);
            pricingHistograms = new PricingHistograms(driver, test);
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC001");
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
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                homePage.VerifyHomePage();
                Thread.Sleep(10000);
                seachPage.VerifyAndEditMoreOptionsInSearchCriteria("Ad Type");
                homePage.VerifyHomePage();
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyPricingHistogramsPage(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Pricing Histograms page.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifyChannelTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify Channel tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                pricingHistograms.VerifyChannel_RetailerTab();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifyChannelTabAnyChartHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify Channel tab any Chart Help icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                pricingHistograms.VerifyChannel_RetailerTab();
                pricingHistograms.VerifyPricingHistogramsPageHelpIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyChannelTabAnyChartWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify Channel tab any Chart Window icon and Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                pricingHistograms.VerifyChannel_RetailerTab();
                Thread.Sleep(10000);
                pricingHistograms.VerifyWindowsIconOnChannelTabCharts("", "Show by #");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab();
                Thread.Sleep(10000);
                pricingHistograms.VerifyWindowsIconOnChannelTabCharts("", "Show by %");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyChannelTabAnyChartExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify Channel tab any Chart Export icon and Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                pricingHistograms.VerifyChannel_RetailerTab();
                Thread.Sleep(10000);
                string chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.png");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab();
                Thread.Sleep(10000);
                chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.jpeg");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab();
                Thread.Sleep(10000);
                chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.pdf");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab();
                Thread.Sleep(10000);
                chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.xlsx");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab();
                Thread.Sleep(10000);
                chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.pptx");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab();
                Thread.Sleep(10000);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_VerifyRetailerTab(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify Retailer tab.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(15000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_VerifyRetailerTabAnyChartHelpIcon(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009-Verify Retailer tab any Chart Help icon.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(15000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPageHelpIcon();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC009");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyRetailerTabAnyChartWindowIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify Retailer tab any Chart Window icon and Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(15000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                pricingHistograms.VerifyWindowsIconOnChannelTabCharts("", "Show by #");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                pricingHistograms.VerifyWindowsIconOnChannelTabCharts("", "Show by %");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyRetailerTabAnyChartExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify Retailer tab any Chart Export icon and Options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(15000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                string chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download PNG");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.png");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download JPG");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.jpeg");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download PDF");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.pdf");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download EXCEL");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.xlsx");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                chartName = pricingHistograms.VerifyExportIconOnChannelTabCharts("", "Download PowerPoint");
                homePage.VerifyFileDownloadedOrNotOnScreen(chartName, "*.pptx");
                Thread.Sleep(10000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyWhenClickAnyDataPointToViewProductImages(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify when Click any data point to view product Images.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(100000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(100000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                pricingHistograms.VerifyDataPointPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC013_VerifyDataPointPopupRadioOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC013-Verify data point popup radio option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(15000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                pricingHistograms.VerifyDataPointPopup();
                Thread.Sleep(10000);
                manufacturerComparison.VerifyRadioButtonsOnDetailDataSection();
                Thread.Sleep(10000);
                manufacturerComparison.VerifyRadioButtonsOnDetailDataSection("Promoted Product Images");
                Thread.Sleep(10000);
                manufacturerComparison.VerifyRadioButtonsOnDetailDataSection("Page Images");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC013");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_VerifyDataPointPopupExportIconAndOptions(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014-Verify data point popup Export icon and options.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(15000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                pricingHistograms.VerifyDataPointPopup();
                Thread.Sleep(5000);
                manufacturerComparison.VerifyExportAsExcelOptionFromDetailDataSection();
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_Detail_Report_Promoted_Product", "*.xlsx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC014");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifyDataPointPopupPageNavigationOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify data point popup Page Navigation option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(15000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                pricingHistograms.VerifyDataPointPopup();
                Thread.Sleep(5000);
                adDetails.VerifyPaginationFunctionality("Next");
                Thread.Sleep(5000);
                adDetails.VerifyPaginationFunctionality("Last");
                Thread.Sleep(5000);
                adDetails.VerifyPaginationFunctionality("Previous");
                Thread.Sleep(5000);
                adDetails.VerifyPaginationFunctionality("First");
                Thread.Sleep(5000);
                adDetails.VerifyPaginationFunctionality("3");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifyDataPointPopupShowPerPageOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify data point popup Show per Page option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(15000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                pricingHistograms.VerifyDataPointPopup();
                Thread.Sleep(5000);
                manufacturerComparison.VerifyRadioButtonsOnDetailDataSection("Page Images");
                Thread.Sleep(5000);
                manufacturerComparison.VerifyItemsPerPageFunctionality("40");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_VerifyDataPointPopupCancelButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017-Verify data point popup Cancel button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                Thread.Sleep(20000);
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Target");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Pricing & Promotions");
                Thread.Sleep(10000);
                pricingHistograms.VerifyPricingHistogramsPage();
                Thread.Sleep(15000);
                pricingHistograms.VerifyChannel_RetailerTab("Retailer");
                Thread.Sleep(10000);
                pricingHistograms.VerifyDataPointPopup(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite015_PricingHistograms_TC017");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
