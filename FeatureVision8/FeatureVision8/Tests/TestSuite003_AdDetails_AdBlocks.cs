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
    public class TestSuite003_AdDetails_AdBlocks : Base
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
            Results.WriteTestSuiteHeading(typeof(TestSuite003_AdDetails_AdBlocks).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite003_AdDetails_AdBlocks).Name);

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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyMedlibSearchInPromoSearchArea(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify Medlib search in Promo Search Area.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                seachPage.VerifyAndEditMoreOptionsInSearchCriteria();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC002");
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
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select Visible Ads");
                Thread.Sleep(2000);
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.VerifySelectionOfRecords(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC003_4");
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
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                seachPage.VerifySearchPageAndSelectCategory("Date", new string[] { "Dynamic Date Range", "Week Starting", "Nielsen Week Ending", "Month", "Ad Date" }, "Month");
                seachPage.VerifySearchPageAndSelectCategory("Month", null, "February - 2020", "Run Report");
                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select Visible Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Deselect Visible Ads");
                homePage.VerifyHomePage();
                adDetails.VerifySelectionOfRecords(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC005");
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
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select All Ads");
                homePage.VerifyHomePage();
                adDetails.VerifySelectionOfRecords(true, false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC006");
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
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC007");
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
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Tiles");
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC008_9");
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
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocksTC010");
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
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Sort By");
                adDetails.VerifyChooseMultipleColumnsToSortByPopup(true, false, "Cancel");
                adDetails.VerifyChooseMultipleColumnsToSortByPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC011_13");
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
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC012");
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
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC014_15");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifyWhenUserClickOnSavePromoSearchButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify when user click on Save Promo Search button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Save Search", "Save Search");
                adDetails.VerifySaveSearchPopup();
                string searchName = adDetails.EditSaveSearchPopup("Random");
                adDetails.ClickButtonInPopup("Save Promo Search");
                homePage.VerifyAlertPopupMessageAndClickButton("Promo Search \"" + searchName + "\" saved successfully.", "Okay, ");
                adDetails.VerifySaveSearchPopup(false);
                homePage.VerifyHomePage();
                homePage.VerifyHomeScreenInDetail(searchName);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_VerifyWhenUserClickOnCancelButtonInSaveSearchPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017-Verify When user click on Cancel button in Save Search pop up.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Save Search", "Save Search");
                adDetails.VerifySaveSearchPopup();
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifySaveSearchPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC017");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC018_21_VerifyWhenUserClicksOnExcelOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC018_21-Verify when user click on Excel option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Excel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, false, true);
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC018_21");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC019_20_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnDownloadReportButtonInOptionForCreatingYourAdBlocksReportPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC019_20-Verify when user select few record on grid and click on Download report button in Option for creating your Ad Blocks Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Excel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, true);
                adDetails.ClickButtonInPopup("Download report");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//*[contains(text(), 'Please wait while export file is being prepared.')]"), "'Preparing for Download' Message did not appear.");
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_Detail_Report_Ad_Blocks_", "*.xlsx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC019_20");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC022_25_VerifyWhenUserClicksOnPDFOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC022_25-Verify when user click on PDF option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "PDF");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, false, false);
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC022_25");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC023_24_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnDownloadReportButtonInOptionForCreatingYourPromotedProductsReportPopupForPDFOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC023_24-Verify when user select few record on grid and click on Download report button in Option for creating your Promoted Products Report popup for PDF option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC023_24");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC026_VerifyWhenUserClicksOnPowerpointOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC026-Verify when user click on Powerpoint option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Power point");
                homePage.VerifyAlertPopupMessageAndClickButton("There are no images selected for PowerPoint report. Please select images using the available checkboxes", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC026");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC027_31_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnPowerpointOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC027_31-Verify when user Selects a few records and then on Powerpoint option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Power point");
                adDetails.VerifyImageReportOptionsPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC027_31");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC028_VerifyWhenUserClicksDownloadReportButtonInImageReportOptionsForSelectedPagesPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC028-Verify when user click on Download Report button in Image Report Options For Selected Pages popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC028");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC029_VerifyWhenUserClicksEmailReportAsAttachmentButtonInImageReportOptionsForSelectedPagesPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC029-Verify when user click on Email Report As Attachment button in Image Report Options For Selected Pages popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC029");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC030_VerifyWhenUserClicksEmailReportAsLinkButtonInImageReportOptionsForSelectedPagesPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC030-Verify when user click on Email Report as Link button in Image Report Options For Selected Pages popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC030");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC032_35_VerifyWhenUserClicksOnEmailOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC032_35-Verify when user click on Email option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Email Option");
                adDetails.VerifySendEmailAsPopup();
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifySendEmailAsPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC032_35");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC033_VerifyWhenUserClicksOnEmailAsAttachmentButtonInSendEmailAsPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC033-Verify When User Clicks On Email As Attachment Button In Send Email As Popup");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Email Option");
                adDetails.VerifySendEmailAsPopup();
                adDetails.ClickButtonInPopup("Email As Attachment");
                homePage.VerifyAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
                adDetails.VerifySendEmailAsPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC033");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC034_VerifyWhenUserClicksOnEmailWithDownloadLinkButtonInSendEmailAsPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC034-Verify when user click on Email with Download Link button in Send Email As popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Email Option");
                adDetails.VerifySendEmailAsPopup();
                adDetails.ClickButtonInPopup("Email with Download Link");
                homePage.VerifyAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
                adDetails.VerifySendEmailAsPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC034");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC036_VerifyWhenUserClickOnResetAllSelectionsOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC036-Verify when user click on Reset All Selections option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Reset All Selections");
                homePage.VerifyAlertPopupMessageAndClickButton("You have not made any changes to default selection.", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC036");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC037_38_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnResetAllSelectionsOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC037_38-Verify when user click on Reset All Selections option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC037_38");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC039_VerifyWhenUserClickOnCancelButtonOnResetAllSelectionsPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC039-Verify when user click on Cancel Button on Reset All Selections Popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC039");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC040_41_VerifyAdBlocksTabAndAgGrid(String Bname)
        {
            TestFixtureSetUp(Bname, "TC040_41-Verify Ad Blocks Tab and AgGrid.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC040_41");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC045_VerifySortOptionForAvailableColumns(String Bname)
        {
            TestFixtureSetUp(Bname, "TC045-Verify Sort option for available columns");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.VerifySortingOnMultipleColumns(9);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC045");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC046_VerifyWhenUserClickOnAnyRecordFromTabularView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC046-Verify when User click on any Record from Tabular view.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC046");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC047_VerifyWhenUserClickOnViewAdFromTilesView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC047-Verify when User click on View Ad from Tiles view.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC047");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC048_VerifyWhenUserClickOnDetailFromThumbnailView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC048-Verify when User click on Detail from Thumbnail view.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC048");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC049_50_VerifyWhenUserClickOnDownloadOnViewAdPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC049_50-Verify when User click on Download on View Ad Popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC049_50");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC051_VerifyPaginationFunctionality(String Bname)
        {
            TestFixtureSetUp(Bname, "TC051-Verify Pagination functionality.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.VerifyPaginationFunctionality("4");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC051");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC052_VerifyItemsPerPageFunctionality(String Bname)
        {
            TestFixtureSetUp(Bname, "TC052-Verify Items Per Page functionality.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Ad Blocks");
                homePage.VerifyHomePage();
                adDetails.VerifyItemsPerPageFunctionality("20");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite003_AdDetails_AdBlocks_TC052");
                throw;
            }
            driver.Quit();
        }


        #endregion
    }
}
