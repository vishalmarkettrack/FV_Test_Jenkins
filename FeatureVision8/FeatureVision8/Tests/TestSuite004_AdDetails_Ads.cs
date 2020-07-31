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
    public class TestSuite004_AdDetails_Ads : Base
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        AdDetails adDetails;
        Search seachPage;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite004_AdDetails_Ads).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite004_AdDetails_Ads).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyTheAdsTabAfterUserSuccessfullyLogsIn(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify the Ads tab after user successfully logs in.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_4_VerifyWhenUserClicksOnSelectVisibleAdsOptionFromMultiSelect(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003_4-Verify when user click on Select Visible Ads option from Multi-select.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select Visible Ads");
                Thread.Sleep(2000);
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.VerifySelectionOfRecords(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC003_4");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifyWhenUserClicksOnDeselectVisibleAdsOptionFromMultiSelect(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify when user click on Deselect Visible Ads option from Multi-select.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "February - 2020", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select Visible Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Deselect Visible Ads");
                homePage.VerifyHomePage();
                adDetails.VerifySelectionOfRecords(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyWhenUserClicksOnSelectAllAdsOptionFromMultiSelect(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify when user click on Select All Ads option from Multi-select.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "February - 2020", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select All Ads");
                homePage.VerifyHomePage();
                adDetails.VerifySelectionOfRecords(true, false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyWhenUserClicksOnDeselectAllAdsOptionFromMultiSelect(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify when user click on Deselect All Ads option from Multi-select.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "February - 2020", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select All Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Deselect All Ads");
                homePage.VerifyHomePage();
                adDetails.VerifySelectionOfRecords(false, false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_9_VerifyWhenUserClicksOnTiles_DefaultView_OptionFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008_9-Verify when user click on Tiles(default view) option from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Tiles");
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC008_9");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyWhenUserClicksOnTableOptionFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify when user click on Table option from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_13_VerifyWhenUserClicksOnSortByOptionFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011_13-Verify when user click on Sort By... option from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Sort By");
                adDetails.VerifyChooseMultipleColumnsToSortByPopup(true, false, "Cancel");
                adDetails.VerifyChooseMultipleColumnsToSortByPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC011_13");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyWhenUserClicksOnSortButtonInChooseMultipleColumnsToSortByPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify when user click on Sort Report button in Choose Multiple Columns to Sort by... popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Sort By");
                adDetails.VerifyChooseMultipleColumnsToSortByPopup(true, true);
                string[] sortByFieldsList = adDetails.CaptureSortByFieldsInOrder();
                adDetails.VerifyChooseMultipleColumnsToSortByPopup(true, false, "Cancel");
                adDetails.VerifyChooseMultipleColumnsToSortByPopup(false);
                adDetails.VerifySortByFieldsOrderInMadLibSearchGrid(sortByFieldsList);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_15_VerifyWhenUserClicksOnShowSelectedAdsOrShowAllAdsOptionFromView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014_15-Verify when user click on Show Selected Ads or Show All Ads option from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Show Selected Ads");
                adDetails.VerifyShowAllOrShowSelectedOptionFromViewMenu(true, 4);
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Show All Ads");
                adDetails.VerifyShowAllOrShowSelectedOptionFromViewMenu(false, 4);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC014_15");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_17_20_VerifyWhenUserClicksOnExcelOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016_17_20-Verify when user click on Excel option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Excel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, false, true);
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC016_17_20");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC018_19_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnDownloadReportButtonInOptionForCreatingYourAdBlocksReportPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC018_19-Verify when user select few record on grid and click on Download report button in Option for creating your Ad Blocks Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Excel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, true);
                adDetails.ClickButtonInPopup("Download report");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//*[contains(text(), 'Please wait while export file is being prepared.')]"), "'Preparing for Download' Message did not appear.");
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_Detail_Report_Ads_", "*.xlsx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC018_19");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC021_24_VerifyWhenUserClicksOnPDFOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC021_24-Verify when user click on PDF option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "PDF");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, false, false);
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC021_24");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC022_23_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnDownloadReportButtonInOptionForCreatingYourPromotedProductsReportPopupForPDFOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC022_23-Verify when user select few record on grid and click on Download report button in Option for creating your Promoted Products Report popup for PDF option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "PDF");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, true, false);
                adDetails.ClickButtonInPopup("Download report");
                Thread.Sleep(10000);
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//*[contains(text(), 'Please wait while export file is being prepared.')]"), "'Preparing for Download' Message did not appear.");
                Thread.Sleep(10000);
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_Detail_Report_", "*.pdf");
                adDetails.VerifyReportReadyPopupMessageAndClickButton("If the download of your Product Detail Report has NOT already occured,", "Close");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC022_23");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC025_VerifyWhenUserClicksOnPowerpointOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC025-Verify when user click on Powerpoint option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Power point");
                homePage.VerifyAlertPopupMessageAndClickButton("There are no images selected for PowerPoint report. Please select images using the available checkboxes", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC025");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC026_30_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnPowerpointOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC026_30-Verify when user Selects a few records and then on Powerpoint option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Power point");
                adDetails.VerifyImageReportOptionsPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC026_30");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC027_VerifyWhenUserClicksDownloadReportButtonInImageReportOptionsForSelectedPagesPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC027-Verify when user click on Download Report button in Image Report Options For Selected Pages popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Power point");
                adDetails.VerifyImageReportOptionsPopup();
                adDetails.ClickButtonInPopup("Download report");
                Thread.Sleep(10000);
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//*[contains(text(), 'Please wait while export file is being prepared.')]"), "'Preparing for Download' Message did not appear.");
                Thread.Sleep(30000);
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator_Promotions_Intel_Image_Report", "*.pptx");
                adDetails.VerifyReportReadyPopupMessageAndClickButton("If the download of your Product Detail Report has NOT already occured,", "Close");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC027");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC028_VerifyWhenUserClicksEmailReportAsAttachmentButtonInImageReportOptionsForSelectedPagesPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC028-Verify when user click on Email Report As Attachment button in Image Report Options For Selected Pages popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Power point");
                adDetails.VerifyImageReportOptionsPopup();
                adDetails.ClickButtonInPopup("Email Report as Attachment");
                Thread.Sleep(10000);
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//*[contains(text(), 'Please wait while export file is being prepared.')]"), "'Preparing for Download' Message did not appear.");
                Thread.Sleep(10000);
                adDetails.VerifyImageReportOptionsPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC028");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC029_VerifyWhenUserClicksEmailReportAsLinkButtonInImageReportOptionsForSelectedPagesPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC029-Verify when user click on Email Report as Link button in Image Report Options For Selected Pages popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Power point");
                adDetails.VerifyImageReportOptionsPopup();
                adDetails.ClickButtonInPopup("Email Report as Link");
                Thread.Sleep(10000);
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//*[contains(text(), 'Please wait while export file is being prepared.')]"), "'Preparing for Download' Message did not appear.");
                Thread.Sleep(10000);
                adDetails.VerifyImageReportOptionsPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC029");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC031_34_VerifyWhenUserClicksOnEmailOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC031_34-Verify when user click on Email option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Email Option");
                adDetails.VerifySendEmailAsPopup();
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifySendEmailAsPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC031_34");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC032_VerifyWhenUserClicksOnEmailAsAttachmentButtonInSendEmailAsPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC032-Verify When User Clicks On Email As Attachment Button In Send Email As Popup");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Email Option");
                adDetails.VerifySendEmailAsPopup(true, true);
                adDetails.ClickButtonInPopup("Email As Attachment");
                homePage.VerifyAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
                adDetails.VerifySendEmailAsPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC032");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC033_VerifyWhenUserClicksOnEmailWithDownloadLinkButtonInSendEmailAsPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC033-Verify when user click on Email with Download Link button in Send Email As popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Email Option");
                adDetails.VerifySendEmailAsPopup(true, true);
                adDetails.ClickButtonInPopup("Email with Download Link");
                homePage.VerifyAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
                adDetails.VerifySendEmailAsPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC033");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC035_VerifyWhenUserClickOnResetAllSelectionsOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC035-Verify when user click on Reset All Selections option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Reset All Selections");
                homePage.VerifyAlertPopupMessageAndClickButton("You have not made any changes to default selection.", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC035");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC036_37_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnResetAllSelectionsOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC036_37-Verify when user click on Reset All Selections option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-Select", "Select Visible Ads");
                homePage.VerifyHomePage();
                adDetails.VerifySelectionOfRecords(true);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Reset All Selections");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to reset current result?", "Ok");
                homePage.VerifyHomePage();
                adDetails.VerifySelectionOfRecords(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC036_37");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC038_VerifyWhenUserClickOnCancelButtonOnResetAllSelectionsPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC038-Verify when user click on Cancel Button on Reset All Selections Popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-Select", "Select Visible Ads");
                homePage.VerifyHomePage();
                adDetails.VerifySelectionOfRecords(true);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Reset All Selections");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to reset current result?", "Cancel");
                homePage.VerifyHomePage();
                adDetails.VerifySelectionOfRecords(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC038");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC039_VerifyAdBlocksTabAndAgGrid(String Bname)
        {
            TestFixtureSetUp(Bname, "TC039-Verify Ad Blocks Tab and AgGrid.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC039");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC042_VerifySortOptionForAvailableColumns(String Bname)
        {
            TestFixtureSetUp(Bname, "TC042-Verify Sort option for available columns");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.VerifySortingOnMultipleColumns(9);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC042");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC043_VerifyWhenUserClickOnAnyRecordFromTableView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC043-Verify when User click on any Record from Table view.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC043");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC044_VerifyWhenUserClickOnViewAdFromThumbnailsView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC044-Verify when User click on View Ad from Thumbnails view.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC044");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC045_VerifyWhenUserClickOnDetailFromTilesView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC045-Verify when User click on Detail from Tiles view.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC045");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC046_VerifyWhenUserClickOnDownloadOnViewAdPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC046-Verify when User click on Download on View Ad Popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup();
                adDetails.VerifyDownloadFunctionalityInViewAdPopup("Download Current Page");
                homePage.VerifyFileDownloadedOrNotOnScreen("", "*.pdf");
                adDetails.VerifyDownloadFunctionalityInViewAdPopup("Download Selected Page");
                homePage.VerifyFileDownloadedOrNotOnScreen("", "*.pdf");
                adDetails.VerifyDownloadFunctionalityInViewAdPopup("Download Full Ad");
                homePage.VerifyFileDownloadedOrNotOnScreen("", "*.pdf");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC046");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC047_VerifyPaginationFunctionality(String Bname)
        {
            TestFixtureSetUp(Bname, "TC047-Verify Pagination functionality.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Dynamic Date Range");
                seachPage.VerifySearchPageAndSelectCategory("Dynamic Date Range", null, "Month - Last 12 Calendar");
                seachPage.VerifySearchPageAndSelectCategory("Dynamic Date Range", null, "Month - Last 12 Calendar - Prior Year", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.VerifyPaginationFunctionality("2");
                homePage.VerifyHomePage();
                adDetails.VerifyPaginationFunctionality("Previous");
                homePage.VerifyHomePage();
                adDetails.VerifyPaginationFunctionality("Next");
                homePage.VerifyHomePage();
                adDetails.VerifyPaginationFunctionality("First");
                homePage.VerifyHomePage();
                adDetails.VerifyPaginationFunctionality("Last");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC047");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC048_VerifyItemsPerPageFunctionality(String Bname)
        {
            TestFixtureSetUp(Bname, "TC048-Verify Items Per Page functionality.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ads");
                homePage.VerifyHomePage();
                adDetails.VerifyItemsPerPageFunctionality("20");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite004_AdDetails_Ads_TC048");
                throw;
            }
            driver.Quit();
        }



        #endregion
    }
}
