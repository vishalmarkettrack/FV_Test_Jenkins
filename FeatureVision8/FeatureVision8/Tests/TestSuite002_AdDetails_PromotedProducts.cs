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
    public class TestSuite002_AdDetails_PromotedProducts : Base
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
            Results.WriteTestSuiteHeading(typeof(TestSuite002_AdDetails_PromotedProducts).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite002_AdDetails_PromotedProducts).Name);

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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyPromotedProductsTabOnceUserSuccessfullyLogsIn(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify Promoted Products tab once user successfully logs in.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomeScreenInDetail();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC002");
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
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.VerifySelectionOfRecords(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC003_4");
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
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select Visible Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Deselect Visible Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.VerifySelectionOfRecords(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC005");
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
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select All Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.VerifySelectionOfRecords(true, false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC006");
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
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Select All Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Multi-select", "Deselect All Ads");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.VerifySelectionOfRecords(false, false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC007");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC008_9");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyWhenUserClicksOnCustomizeColumnsOptionFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify when user click on Customize Columns option from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_VerifyWhenUserAddFieldsToDisplayInCustomizeYourReportpopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014-Verify when user Add Fields to Display in Customize your Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display");
                string[] fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                adDetails.ClickButtonInPopup("Apply Selections");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                string[] fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC014");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifyWhenUserRemoveFieldsFromDisplayInCustomizeYourReportpopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify when user Remove Fields to Display in Customize your Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();
                adDetails.DragFieldsInCustomizeYourReportPopup("All Available Fields");
                string[] fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                adDetails.ClickButtonInPopup("Apply Selections");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                string[] fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifyWhenUserChangesOrderFromDisplayInCustomizeYourReportpopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify when user Change order from Fields to Display box in Customize your Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display", false);
                string[] fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                adDetails.ClickButtonInPopup("Apply Selections");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                string[] fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_18_VerifyWhenUserClicksOnEditIconFromDisplayInCustomizeYourReportpopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017_18-Verify when user click on Edit icon from Fields to Display box in Customize your Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();
                string fieldName = adDetails.EditFieldNameFromFieldsToDisplayBoxInCustomizeYourReportPopup();
                string[] fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                adDetails.ClickButtonInPopup("Apply Selections");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                string[] fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, "Current Template");
                adDetails.EditFieldNameFromFieldsToDisplayBoxInCustomizeYourReportPopup(fieldName, true);
                fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                adDetails.ClickButtonInPopup("Apply Selections");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC017_18");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_13_19_VerifyWhenUserSavesTemplateAsPrivateAndActiveTemplateInDisplayInCustomizeYourReportpopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012_13_19-Verify when user Save template as Private and active template in Customize your Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display", false);
                string templateName1 = adDetails.EnterTemplateNameAndSelectRadioButton("", false);
                adDetails.ClickButtonInPopup("Save & Apply");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName1);
                adDetails.SelectTempleteInCustomizeYourReportPopup("Default Template");
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display", false);
                string[] fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                string templateName2 = adDetails.EnterTemplateNameAndSelectRadioButton("", false, false);
                adDetails.ClickButtonInPopup("Save & Apply");
                homePage.VerifyAlertPopupMessageAndClickButton("Do you want to make \"" + templateName2 + "\" as active template?", "Yes");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                string[] fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName2);
                adDetails.SelectTempleteInCustomizeYourReportPopup(templateName1);
                fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                adDetails.ClickButtonInPopup("Save & Apply");
                homePage.VerifyAlertPopupMessageAndClickButton("Do you want to make \"" + templateName1 + "\" as active template?", "Yes");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName1);
                adDetails.ClickButtonInPopup("Delete");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to delete the selected template?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("Template Deleted Successfully", "Okay, Got It");
                Thread.Sleep(5000);
                adDetails.SelectTempleteInCustomizeYourReportPopup(templateName2);
                adDetails.ClickButtonInPopup("Delete");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to delete the selected template?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("Template Deleted Successfully", "Okay, Got It");
                adDetails.VerifyCustomizeYourReportPopup();
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifyCustomizeYourReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC012_13_19");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC020_VerifyWhenUserSavesTemplateAsPrivateAndNotActiveTemplateInDisplayInCustomizeYourReportpopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC020-Verify when user Save template as Private and not active template in Customize your Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display", false);
                string templateName1 = adDetails.EnterTemplateNameAndSelectRadioButton("", false);
                adDetails.ClickButtonInPopup("Save & Apply");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName1);
                adDetails.SelectTempleteInCustomizeYourReportPopup("Default Template");
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display", false);
                string[] fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                string templateName2 = adDetails.EnterTemplateNameAndSelectRadioButton("", false, false);
                adDetails.ClickButtonInPopup("Save & Apply");
                homePage.VerifyAlertPopupMessageAndClickButton("Do you want to make \"" + templateName2 + "\" as active template?", "No");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                string[] fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName1);
                adDetails.ClickButtonInPopup("Delete");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to delete the selected template?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("Template Deleted Successfully", "Okay, Got It");
                Thread.Sleep(5000);
                adDetails.SelectTempleteInCustomizeYourReportPopup(templateName2);
                adDetails.ClickButtonInPopup("Delete");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to delete the selected template?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("Template Deleted Successfully", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC020");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC021_VerifyWhenUserSavesTemplateAsSharedAndActiveTemplateInDisplayInCustomizeYourReportpopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC021-Verify when user Save template as Shared and active template in Customize your Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display", false);
                string templateName1 = adDetails.EnterTemplateNameAndSelectRadioButton("", false);
                adDetails.ClickButtonInPopup("Save & Apply");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName1);
                adDetails.SelectTempleteInCustomizeYourReportPopup("Default Template");
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display", false);
                string[] fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                string templateName2 = adDetails.EnterTemplateNameAndSelectRadioButton("", true, false);
                adDetails.ClickButtonInPopup("Save & Apply");
                homePage.VerifyAlertPopupMessageAndClickButton("Do you want to make \"" + templateName2 + "\" as active template?", "Yes");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                string[] fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName2);
                IWebElement radioEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='radio' and @value='1']"));
                Assert.AreNotEqual(null, radioEle.GetAttribute("checked"), "'Shared' Radio button was not selected.");
                Results.WriteStatus(test, "Pass", "Selected, 'Shared' Radio Button.");
                adDetails.ClickButtonInPopup("Delete");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to delete the selected template?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("Template Deleted Successfully", "Okay, Got It");
                Thread.Sleep(5000);
                adDetails.SelectTempleteInCustomizeYourReportPopup(templateName1);
                adDetails.ClickButtonInPopup("Delete");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to delete the selected template?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("Template Deleted Successfully", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC021");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC022_24_VerifyWhenUserSavesTemplateAsSharedAndNotActiveTemplateInDisplayInCustomizeYourReportpopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC022_24-Verify when user Save template as Shared and not active template in Customize your Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display", false);
                string templateName1 = adDetails.EnterTemplateNameAndSelectRadioButton("", false);
                adDetails.ClickButtonInPopup("Save & Apply");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName1);
                adDetails.SelectTempleteInCustomizeYourReportPopup("Default Template");
                adDetails.DragFieldsInCustomizeYourReportPopup("Fields to Display", false);
                string[] fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                adDetails.ClickButtonInPopup("Save & Apply");
                homePage.VerifyAlertPopupMessageAndClickButton("Please enter template name", "Okay, Got It");
                string templateName2 = adDetails.EnterTemplateNameAndSelectRadioButton("", true, false);
                adDetails.ClickButtonInPopup("Save & Apply");
                homePage.VerifyAlertPopupMessageAndClickButton("Do you want to make \"" + templateName2 + "\" as active template?", "No");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();
                string[] fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName1);
                adDetails.SelectTempleteInCustomizeYourReportPopup(templateName2);
                Thread.Sleep(5000);
                IWebElement radioEle = driver.FindElement(By.XPath("//div[@class='modal-body-filters']//input[@type='radio' and @value='1']"));
                Assert.AreNotEqual(null, radioEle.GetAttribute("checked"), "'Shared' Radio button was not selected.");
                Results.WriteStatus(test, "Pass", "Selected, 'Shared' Radio Button.");
                adDetails.ClickButtonInPopup("Delete");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to delete the selected template?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("Template Deleted Successfully", "Okay, Got It");
                Thread.Sleep(5000);
                adDetails.SelectTempleteInCustomizeYourReportPopup(templateName1);
                adDetails.ClickButtonInPopup("Delete");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to delete the selected template?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("Template Deleted Successfully", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC022_24");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC023_25_VerifyWhenUserSavesTemplateAsSharedAndActiveTemplateInDisplayInCustomizeYourReportpopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC023_25-Verify when user Save template as Shared and active template in Customize your Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup();

                string[] fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                adDetails.DragFieldsInCustomizeYourReportPopup("All Available Fields");
                string[] fieldsDisplayedInBox1 = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                adDetails.ClickButtonInPopup("Reset");
                string[] fieldsDisplayedInBox2 = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedInBox1, true, false);
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedInBox2);

                adDetails.DragFieldsInCustomizeYourReportPopup("All Available Fields");
                fieldsDisplayedInBox = adDetails.CaptureFieldsFromCustomizeYourReportPopup("Fields To Display");
                string templateName1 = adDetails.EnterTemplateNameAndSelectRadioButton("", false);
                adDetails.ClickButtonInPopup("Save & Apply");
                adDetails.VerifyCustomizeYourReportPopup(false);
                homePage.VerifyHomePage();

                string[] fieldsDisplayedinTable = adDetails.CaptureFieldsFromMadlibSearchTableView();
                homePage.CompareStringLists(fieldsDisplayedInBox, fieldsDisplayedinTable);

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Customize Columns");
                adDetails.VerifyCustomizeYourReportPopup(true, false, templateName1);
                adDetails.ClickButtonInPopup("Delete");
                homePage.VerifyAlertPopupMessageAndClickButton("Are you sure you want to delete the selected template?", "Ok");
                homePage.VerifyAlertPopupMessageAndClickButton("Template Deleted Successfully", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC023_25");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC026_28_VerifyWhenUserClicksOnSortByOptionFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC026_28-Verify when user click on Sort By... option from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Sort By");
                adDetails.VerifyChooseMultipleColumnsToSortByPopup(true, false, "Cancel");
                adDetails.VerifyChooseMultipleColumnsToSortByPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC026_28");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC027_VerifyWhenUserClicksOnSortButtonInChooseMultipleColumnsToSortByPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC027-Verify when user click on Sort Report button in Choose Multiple Columns to Sort by... popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC027");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC029_30_VerifyWhenUserClicksOnShowSelectedAdsOrShowAllAdsOptionFromView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC029_30-Verify when user click on Show Selected Ads or Show All Ads option from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC029_30");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC031_32_36_VerifyWhenUserClicksOnExcelOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC031_32_36-Verify when user click on Excel option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Excel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, false, true);
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC031_32_36");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC033_34_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnDownloadReportButtonInOptionForCreatingYourPromotedProductsReportPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC033_34-Verify when user select few record on grid and click on Download report button in Option for creating your Promoted Products Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Excel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, true);
                adDetails.ClickButtonInPopup("Download report");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//*[contains(text(), 'Please wait while export file is being prepared.')]"), "'Preparing for Download' Message did not appear.");
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_Detail_Report_Promoted_Products_", "*.xlsx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC033_34");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC035_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnSendMailWithAttachedReportButtonInOptionForCreatingYourPromotedProductsReportPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC035-Verify when user select few record on grid and click on Send mail with attached report button in Option for creating your Promoted Products Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Excel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, true);
                adDetails.ClickButtonInPopup("Send mail with attached report");
                homePage.VerifyAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC035");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC037_41_VerifyWhenUserClicksOnPDFOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC037_41-Verify when user click on PDF option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "PDF");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, false, false);
                adDetails.ClickButtonInPopup("Cancel");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC037_41");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC038_39_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnDownloadReportButtonInOptionForCreatingYourPromotedProductsReportPopupForPDFOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC038_39-Verify when user select few record on grid and click on Download report button in Option for creating your Promoted Products Report popup for PDF option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "PDF");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, true, false);
                adDetails.ClickButtonInPopup("Download report");
                Thread.Sleep(10000);
                Assert.IsTrue(driver._waitForElementToBeHidden("xpath", "//*[contains(text(), 'Please wait while export file is being prepared.')]"), "'Preparing for Download' Message did not appear.");
                Thread.Sleep(10000);
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_Detail_Report_Promoted_Products_", "*.pdf");
                adDetails.VerifyReportReadyPopupMessageAndClickButton("If the download of your Product Detail Report has NOT already occured,", "Close");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC038_39");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC040_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnSendEmailWithAttachedReportButtonInOptionForCreatingYourPromotedProductsReportPopupForPDFOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC040-Verify when user select few record on grid and click on Send Email With Attached Report button in Option for creating your Promoted Products Report popup for PDF option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "PDF");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(true, true, false);
                adDetails.ClickButtonInPopup("Send mail with attached report");
                homePage.VerifyAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
                adDetails.VerifyOptionForCreatingYourPromotedProductsReportPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC040");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC042_VerifyWhenUserClicksOnPowerpointOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC042-Verify when user click on Powerpoint option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Power point");
                homePage.VerifyAlertPopupMessageAndClickButton("There are no images selected for PowerPoint report. Please select images using the available checkboxes", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC042");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC043_VerifyWhenUserSelectsFewRecordsOnGridAndClicksOnPowerpointOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC043-Verify when user Selects a few records and then on Powerpoint option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                adDetails.SelectRecordsFromMadlibSearchResultGrid(4);
                adDetails.SelectMenuOptionInMadlibSearchView("Export", "Power point");
                adDetails.VerifyImageReportOptionsPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC043");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC044_VerifyNormalFilterOptionForAvailableColumns(String Bname)
        {
            TestFixtureSetUp(Bname, "TC044-Verify Normal Filter option for available columns (i.e. Account, Date)");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();

                string column = adDetails.GetFilterTypeForColumnOrColumnForFilterType("Normal");
                adDetails.VerifyNormalFilterFunctionality(column);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC044");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC047_VerifySortOptionForAvailableColumns(String Bname)
        {
            TestFixtureSetUp(Bname, "TC047-Verify Sort option for available columns");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.VerifySortingOnMultipleColumns(9);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC047");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC048_VerifyWhenUserClickOnAnyRecordFromTabularView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC048-Verify when User click on any Record from Tabular view.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();

                adDetails.SelectMenuOptionInMadlibSearchView("Views", "Table");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC048");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC049_VerifyWhenUserClickOnViewAdFromTilesView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC049-Verify when User click on View Ad from Tiles view.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC049");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC050_VerifyWhenUserClickOnDetailFromThumbnailView(String Bname)
        {
            TestFixtureSetUp(Bname, "TC050-Verify when User click on Detail from Thumbnail view.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.OpenViewAdOrDetailsPopup(true);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC050");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC051_VerifyWhenUserClickOnDownloadOnViewAdPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC051-Verify when User click on Download on View Ad Popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC051");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC052_VerifyPaginationFunctionality(String Bname)
        {
            TestFixtureSetUp(Bname, "TC052-Verify Pagination functionality.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC052");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC053_VerifyItemsPerPageFunctionality(String Bname)
        {
            TestFixtureSetUp(Bname, "TC053-Verify Items Per Page functionality.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                adDetails.SelectTabInMadlibSearchView("Promoted Products");
                homePage.VerifyHomePage();
                adDetails.VerifyItemsPerPageFunctionality("20");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite002_AdDetails_PromotedProducts_TC053");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }


}











