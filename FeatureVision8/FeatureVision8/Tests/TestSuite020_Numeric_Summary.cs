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
    public class TestSuite020_Numeric_Summary : Base
    {
        string clientName = "Procter & Gamble";

        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        SummarySection summarySection;
        ExecutiveDashboard executiveDashboard;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite020_Numeric_Summary).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite020_Numeric_Summary).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            executiveDashboard = new ExecutiveDashboard(driver, test);
            summarySection = new SummarySection(driver, test);

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
                loginPage.loginAndVerifyHomePageWithClient(clientName);
                homePage.VerifyHomeScreenInDetail();

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifySummaryOptionsPopupWindow(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify Summary Options popup window.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyNumericSummaryTabSection(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify Numeric Summary tab Section.");
            try
            {
                loginPage.loginAndVerifyHomePageWithClient(clientName);

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_SelectFormatOptionAndRunReport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Select Format Option and Run Report.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();
                summarySection.clickSelectFormatDropdownAndSelectOptionFromList("Retailer");
                summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");

                string[] formatCollections = { "Parent Retailer", "Brand", "Category", "Category Group", "Department", "Manufacturer", "Market", "Product Description", "Segment", "Brand Family" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryGridSection();
                    summarySection.clickEditLinkFromSummaryGridSection("Numeric Summary");
                    summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                    summarySection.verifyNumericSummarySection();
                    summarySection.clickSelectFormatDropdownAndSelectOptionFromList(formatCollections[i]);
                    summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_01_SelectFormatOptionAndRunReport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004_01-Select Format Option and Run Report.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();
                summarySection.clickSelectFormatDropdownAndSelectOptionFromList("Media Type");
                summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");

                string[] formatCollections = { "Page Position", "Brand/Retailer", "Brand/Parent Retailer", "Brand/Category", "Brand/Category Group", "Brand/Department", "Brand/Market", "Category/Retailer", "Category/Parent Retailer" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryGridSection();
                    summarySection.clickEditLinkFromSummaryGridSection("Numeric Summary");
                    summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                    summarySection.verifyNumericSummarySection();
                    summarySection.clickSelectFormatDropdownAndSelectOptionFromList(formatCollections[i]);
                    summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC004_01");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_02_SelectFormatOptionAndRunReport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004_02-Select Format Option and Run Report.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();
                summarySection.clickSelectFormatDropdownAndSelectOptionFromList("Category/Market");
                summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");

                string[] formatCollections = { "Category Group/Retailer", "Category Group/Parent Retailer", "Category Group/Market", "Department/Retailer", "Department/Parent Retailer", "Department/Market", "Manufacturer/Retailer", "Manufacturer/Parent Retailer", "Manufacturer/Category", "Manufacturer/Category Group" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryGridSection();
                    summarySection.clickEditLinkFromSummaryGridSection("Numeric Summary");
                    summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                    summarySection.verifyNumericSummarySection();
                    summarySection.clickSelectFormatDropdownAndSelectOptionFromList(formatCollections[i]);
                    summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC004_02");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_03_SelectFormatOptionAndRunReport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004_03-Select Format Option and Run Report.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();
                summarySection.clickSelectFormatDropdownAndSelectOptionFromList("Manufacturer/Department");
                summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");

                string[] formatCollections = { "Manufacturer/Market", "Retailer/Market", "Retailer/Media Type", "Market/Media Type", "Segment/Retailer", "Segment/Market", "Brand Family/Retailer", "Brand Family/Market", "Brand/Page Position", "Segment/Parent Retailer" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryGridSection();
                    summarySection.clickEditLinkFromSummaryGridSection("Numeric Summary");
                    summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                    summarySection.verifyNumericSummarySection();
                    summarySection.clickSelectFormatDropdownAndSelectOptionFromList(formatCollections[i]);
                    summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC004_03");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_04_SelectFormatOptionAndRunReport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004_04-Select Format Option and Run Report.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();
                summarySection.clickSelectFormatDropdownAndSelectOptionFromList("Segment/Channel");
                summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");

                string[] formatCollections = { "Manufacturer/Page Position", "Category/Page Position", "Segment/Manufacturer", "Manufacturer/Product Description", "Retailer/Product Description", "Brand Family/Channel", "Brand/Channel", "Manufacturer/Brand", "Segment/Brand" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryGridSection();
                    summarySection.clickEditLinkFromSummaryGridSection("Numeric Summary");
                    summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                    summarySection.verifyNumericSummarySection();
                    summarySection.clickSelectFormatDropdownAndSelectOptionFromList(formatCollections[i]);
                    summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC004_04");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_SelectFormatOptionAndRunReport_Target_Durable(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Select Format Option and Run Report (Target --> Durable).");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();
                summarySection.clickSelectFormatDropdownAndSelectOptionFromList("Retailer");
                summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");

                string[] formatCollections = { "Retailer Group", "Brand", "Category", "Department", "Division", "Manufacturer", "Market", "Pyramid", "Media Type", "Page Position" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryGridSection();
                    summarySection.clickEditLinkFromSummaryGridSection("Numeric Summary");
                    summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                    summarySection.verifyNumericSummarySection();
                    summarySection.clickSelectFormatDropdownAndSelectOptionFromList(formatCollections[i]);
                    summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_01_SelectFormatOptionAndRunReport_Target_Durable(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005_01-Select Format Option and Run Report (Target --> Durable).");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();
                summarySection.clickSelectFormatDropdownAndSelectOptionFromList("Brand/Retailer");
                summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");

                string[] formatCollections = { "Brand/Retailer Group", "Brand/Category", "Brand/Department", "Brand/Division", "Brand/Market", "Category/Retailer", "Category/Retailer Group", "Category/Market", "Department/Retailer", "Department/Retailer Group" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryGridSection();
                    summarySection.clickEditLinkFromSummaryGridSection("Numeric Summary");
                    summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                    summarySection.verifyNumericSummarySection();
                    summarySection.clickSelectFormatDropdownAndSelectOptionFromList(formatCollections[i]);
                    summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC005_01");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_02_SelectFormatOptionAndRunReport_Target_Durable(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005_02-Select Format Option and Run Report (Target --> Durable).");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();
                summarySection.clickSelectFormatDropdownAndSelectOptionFromList("Department/Market");
                summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");

                string[] formatCollections = { "Division/Retailer", "Division/Retailer Group", "Division/Market", "Manufacturer/Retailer", "Manufacturer/Retailer Group", "Manufacturer/Category", "Manufacturer/Department", "Manufacturer/Division", "Manufacturer/Market", "Retailer/Market" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryGridSection();
                    summarySection.clickEditLinkFromSummaryGridSection("Numeric Summary");
                    summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                    summarySection.verifyNumericSummarySection();
                    summarySection.clickSelectFormatDropdownAndSelectOptionFromList(formatCollections[i]);
                    summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC005_02");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_03_SelectFormatOptionAndRunReport_Target_Durable(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005_03-Select Format Option and Run Report (Target --> Durable).");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.clickTabFromPromoSeachSection("Summary");

                summarySection.verifySummaryGridSection();
                summarySection.clickEditLinkFromSummaryGridSection();
                summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                summarySection.verifyNumericSummarySection();
                summarySection.clickSelectFormatDropdownAndSelectOptionFromList("Retailer/Media Type");
                summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");

                string[] formatCollections = { "Market/Media Type", "Pyramid/Retailer", "Pyramid/Market", "Brand/Page Position", "Pyramid/Retailer Group", "Manufacturer/Page Position", "Category/Page Position", "Pyramid/Manufacturer", "Manufacturer/Brand", "Pyramid/Brand" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryGridSection();
                    summarySection.clickEditLinkFromSummaryGridSection("Numeric Summary");
                    summarySection.clickTabFromSummaryOptionsPopupWindow("Numeric Summary");
                    summarySection.verifyNumericSummarySection();
                    summarySection.clickSelectFormatDropdownAndSelectOptionFromList(formatCollections[i]);
                    summarySection.clickButtonFromSummaryOptionsPopupWindow("Run Report");
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC005_03");
                throw;
            }
            driver.Quit();
        }






        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_NumericSummary_Wells_Dairy_Client(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Select Format Option and Run Report (Target --> Durable).");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Summary");
                homePage.verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen("Summary", "Numeric Summary");

                summarySection.verifySummaryOptionsPopupWindow("Retailer", "Numeric Summary");
                summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                summarySection.clickPlusIcon();

                string[] formatCollections = { "Parent Retailer", "Brand", "Category", "Category Group", "Department", "Manufacturer", "Market", "Product Description", "Segment" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryOptionsPopupWindow(formatCollections[i], "Numeric Summary");
                    summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                    summarySection.clickPlusIcon();
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_01_NumericSummary_Wells_Dairy_Client(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006_01-Select Format Option and Run Report (Target --> Durable).");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Summary");
                homePage.verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen("Summary", "Numeric Summary");

                summarySection.verifySummaryOptionsPopupWindow("Media Type", "Numeric Summary");
                summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                summarySection.clickPlusIcon();

                string[] formatCollections = { "Page Position", "Brand/Retailer", "Brand/Parent Retailer", "Brand/Category", "Brand/Category Group", "Brand/Department", "Brand/Market", "Category/Retailer", "Category/Parent Retailer" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryOptionsPopupWindow(formatCollections[i], "Numeric Summary");
                    summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                    summarySection.clickPlusIcon();
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC006_01");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_02_NumericSummary_Wells_Dairy_Client(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006_02-Select Format Option and Run Report (Target --> Durable).");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Summary");
                homePage.verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen("Summary", "Numeric Summary");

                summarySection.verifySummaryOptionsPopupWindow("Category/Market", "Numeric Summary");
                summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                summarySection.clickPlusIcon();

                string[] formatCollections = { "Category Group/Retailer", "Category Group/Parent Retailer", "Category Group/Market", "Department/Retailer", "Department/Parent Retailer", "Department/Market", "Manufacturer/Retailer", "Manufacturer/Parent Retailer", "Manufacturer/Category" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryOptionsPopupWindow(formatCollections[i], "Numeric Summary");
                    summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                    summarySection.clickPlusIcon();
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC006_02");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_03_NumericSummary_Wells_Dairy_Client(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006_03-Select Format Option and Run Report (Target --> Durable).");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Summary");
                homePage.verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen("Summary", "Numeric Summary");

                summarySection.verifySummaryOptionsPopupWindow("Manufacturer/Category Group", "Numeric Summary");
                summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                summarySection.clickPlusIcon();

                string[] formatCollections = { "Manufacturer/Department", "Manufacturer/Market", "Retailer/Market", "Retailer/Media Type", "Market/Media Type", "Segment/Retailer", "Segment/Market", "Brand/Page Position", "Segment/Parent Retailer" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryOptionsPopupWindow(formatCollections[i], "Numeric Summary");
                    summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                    summarySection.clickPlusIcon();
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC006_03");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_04_NumericSummary_Wells_Dairy_Client(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006_04-Select Format Option and Run Report (Target --> Durable).");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Summary");
                homePage.verifySubMenuScreen_With_TitleAndDescriptions_AndClickOnScreen("Summary", "Numeric Summary");

                summarySection.verifySummaryOptionsPopupWindow("Segment/Channel", "Numeric Summary");
                summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                summarySection.clickPlusIcon();

                string[] formatCollections = { "Manufacturer/Page Position", "Category/Page Position", "Segment/Manufacturer", "Manufacturer/Product Description", "Retailer/Product Description", "Brand/Channel", "Manufacturer/Brand", "Segment/Brand" };

                for (int i = 0; i < formatCollections.Length; i++)
                {
                    summarySection.verifySummaryOptionsPopupWindow(formatCollections[i], "Numeric Summary");
                    summarySection.clickDisplayReportButtonOnSummaryOptionsSection();
                    summarySection.clickPlusIcon();
                }

            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite020_Numeric_Summary_TC006_04");
                throw;
            }
            driver.Quit();
        }


        #endregion
    }
}
