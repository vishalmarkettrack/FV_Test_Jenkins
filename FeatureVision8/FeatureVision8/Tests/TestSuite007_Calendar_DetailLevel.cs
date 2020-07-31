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
    public class TestSuite007_Calendar_DetailLevel : Base
    {
        #region Private Variables

        private IWebDriver driver;
        Login loginPage;
        Home homePage;
        AdDetails adDetails;
        Search seachPage;
        SavedSearches savedSearches;
        Calendar calendar;

        #endregion

        #region Public Fixture Methods

        public IWebDriver TestFixtureSetUp(string Bname, string testCaseName)
        {
            driver = StartBrowser(Bname);
            Common.CurrentDriver = driver;
            Results.WriteTestSuiteHeading(typeof(TestSuite007_Calendar_DetailLevel).Name);
            starttest(Bname + " - " + testCaseName, typeof(TestSuite007_Calendar_DetailLevel).Name);

            loginPage = new Login(driver, test);
            homePage = new Home(driver, test);
            adDetails = new AdDetails(driver, test);
            seachPage = new Search(driver, test);
            savedSearches = new SavedSearches(driver, test);
            calendar = new Calendar(driver, test);
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
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC001");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC002_VerifyCalenderScreen(String Bname)
        {
            TestFixtureSetUp(Bname, "TC002-Verify calender screen.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomeScreenInDetail();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC002");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC003_VerifyWhenUserClickOnViewsButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC003-Verify when user click on Views button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC003");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC004_VerifyWhenUserMouseOverOnCalendarDataByFromViewsForUSClient(String Bname)
        {
            TestFixtureSetUp(Bname, "TC004-Verify when user mouse over on Calendar Data by from Views for US client.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC004");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC005_VerifyWhenUserMouseOverOnCalendarDataByFromViewsForCanadaClient(String Bname)
        {
            TestFixtureSetUp(Bname, "TC005-Verify when user mouse over on Calendar Data by from Views for Canada client.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Cascades Canada");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen("Cascades Canada");
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "", "Cascades Canada");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC005");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC006_VerifyWhenUserMouseOverOnCalendarDataByFromViewsForAustraliaClient(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify when user mouse over on Calendar Data by from Views for Australia client.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Metcash - Australia");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen("Metcash - Australia");
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "", "Metcash - Australia");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC006");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC007_VerifyWhenUserClicksOnCalendarDataBy_RetailerFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC007-Verify when user click on Calendar Data by > Retailer from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Retailer");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Retailer", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC007");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC008_VerifyWhenUserClicksOnCalendarDataBy_MarketFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC008-Verify when user click on Calendar Data by > Market from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Market");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Market", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC008");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC009_VerifyWhenUserClicksOnCalendarDataBy_AdTypeFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC009-Verify when user click on Calendar Data by > AdType from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Ad Type");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Ad Type", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC009");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC010_VerifyWhenUserClicksOnCalendarDataBy_ManufacturerFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC010-Verify when user click on Calendar Data by > Manufacturer from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Manufacturer");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Manufacturer", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC010");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC011_VerifyWhenUserClicksOnCalendarDataBy_DepartmentFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC011-Verify when user click on Calendar Data by > Department from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Department");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Department", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC011");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC012_VerifyWhenUserClicksOnCalendarDataBy_CategoryGroupFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC012-Verify when user click on Calendar Data by > Category Group from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Category Group");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Category Group", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC012");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC013_VerifyWhenUserClicksOnCalendarDataBy_CategoryFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC013-Verify when user click on Calendar Data by > Category from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Category");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Category", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC013");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC014_VerifyWhenUserClicksOnCalendarDataBy_BrandFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC014-Verify when user click on Calendar Data by > Brand from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Brand");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Brand", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC014");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC015_VerifyWhenUserClicksOnCalendarDataBy_PagePositionFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC015-Verify when user click on Calendar Data by > Page Position from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Page Position");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Page Position", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC015");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC016_VerifyWhenUserClicksOnCalendarDataBy_MediaTypeFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC016-Verify when user click on Calendar Data by > Media Type from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Media Type");
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Media Type", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC016");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC017_VerifyWhenUserMouseOverOnCalendarBasedOnFromViewsForUSClient(String Bname)
        {
            TestFixtureSetUp(Bname, "TC017-Verify when user mouse over on Calendar Based On from Views for US client.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC017");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC018_VerifyWhenUserMouseOverOnCalendarBasedOnFromViewsForCanadaClient(String Bname)
        {
            TestFixtureSetUp(Bname, "TC018-Verify when user mouse over on Calendar Based On from Views for Canada client.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Cascades Canada");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen("Cascades Canada");
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "", "Cascades Canada");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC018");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC019_VerifyWhenUserMouseOverOnCalendarBasedOnFromViewsForAustraliaClient(String Bname)
        {
            TestFixtureSetUp(Bname, "TC006-Verify when user mouse over on Calendar Based On from Views for Australia client.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Metcash - Australia");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen("Metcash - Australia");
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "", "Metcash - Australia");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC019");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC020_VerifyWhenUserClickOnCalendarBasedOn_AdDateFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC020-Verify when user click on Calendar Based On > Ad Date from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen("Procter & Gamble");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Category", "Procter & Gamble");
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Ad Date", "Procter & Gamble");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Category", "Ad Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC020");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC021_VerifyWhenUserClickOnCalendarBasedOn_StartDateFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC021-Verify when user click on Calendar Based On > Start Date from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen("Procter & Gamble");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Category", "Procter & Gamble");
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Sale Start Date", "Procter & Gamble");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Category", "Sale Start Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC021");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC022_VerifyWhenUserClickOnCalendarBasedOn_EndDateFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC022-Verify when user click on Calendar Based On > End Date from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen("Procter & Gamble");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Data By", "Category", "Procter & Gamble");
                calendar.SelectMenuOptionInCalendarView("Views", "Calendar Based On", "Sale End Date", "Procter & Gamble");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid("Category", "Sale End Date");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC022");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC023_VerifyWhenUserClickOnNormalFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC023-Verify when user click on Normal from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen("Procter & Gamble");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.SelectMenuOptionInCalendarView("Views", "Normal", "", "Procter & Gamble");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC023");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC024_VerifyWhenUserClickOnSideBySideFromViews(String Bname)
        {
            TestFixtureSetUp(Bname, "TC024-Verify when user click on Side By Side from Views.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen("Procter & Gamble");
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.SelectMenuOptionInCalendarView("Views", "Side By Side", "", "Procter & Gamble");
                homePage.VerifyHomePage();
                calendar.VerifyCalendarViewGrid();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC024");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC025_26_VerifyWhenUserClickOnExcelOptionFromExport(String Bname)
        {
            TestFixtureSetUp(Bname, "TC025_26-Verify when user click on Excel option from Export.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                Thread.Sleep(30000);
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC025_26");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC027_37_VerifyWhenUserSelectDoNotIncludeHyperlinksOptionAndClickOnDownloadReportButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC027_37-Verify when user select Do not include hyperlinks option and click on Download Report button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.CaptureDataFromCalendarViewGrid();
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Do not include hyperlink");
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Download Report");
                Thread.Sleep(30000);
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_DetailCalendarReport", "*.xlsx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC027_37");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC028_35_36_VerifyWhenUserSelectDoNotIncludeHyperlinksOptionAndClickOnSendMailWithAttachedReportButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC028_35_36-Verify when user select Do not include hyperlinks option and click on Send Mail With Attached Report button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Do not include hyperlink");
                calendar.EnterReportNameAndEmailSubjectLineInExportExcelPopup("Random", "Random");
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Send Email With Attached Report");
                homePage.VerifyEmailAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC028_35_36");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC029_VerifyWhenUserSelectSaveTheProductDetailRecordsOptionAndClickOnDownloadReportButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC029-Verify when user select Save the Product Detail Records option and click on Download Report button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Save the Product Detail Records");
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Download Report");
                Thread.Sleep(30000);
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_DetailCalendarReport", "*.xlsx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC029");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC030_VerifyWhenUserSelectSaveTheProductDetailRecordsOptionAndClickOnSendMailWithAttachedReportButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC030-Verify when user select Save The Product Detail Records option and click on Send Mail With Attached Report button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Save The Product Detail Records");
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Send Email With Attached Report");
                homePage.VerifyEmailAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC030");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC031_VerifyWhenUserSelectDoNotDisplayQueryParametersOptionAndClickOnDownloadReportButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC031-Verify when user select Do not Display Query Parameters option and click on Download Report button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Do not Display Query Parameters");
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Download Report");
                Thread.Sleep(30000);
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_DetailCalendarReport", "*.xlsx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC031");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC032_VerifyWhenUserSelectDoNotDisplayQueryParametersOptionAndClickOnSendMailWithAttachedReportButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC032-Verify when user select Do not display query parameters option and click on Send Mail With Attached Report button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Do not display query Parameters");
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Send Email With Attached Report");
                homePage.VerifyEmailAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC032");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC033_VerifyWhenUserSelectDeliverReportAsWinZipFileOptionAndClickOnDownloadReportButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC033-Verify when user select Deliver report as WinZip file option and click on Download Report button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Deliver report as WinZip file");
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Download Report");
                Thread.Sleep(30000);
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_DetailCalendarReport", "*.zip");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC033");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC034_VerifyWhenUserSelectDeliverReportAsWinZipFileOptionAndClickOnSendMailWithAttachedReportButton(String Bname)
        {
            TestFixtureSetUp(Bname, "TC034-Verify when user select Deliver report as WinZip file option and click on Send Mail With Attached Report button.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Deliver report as WinZip file");
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Send Email With Attached Report");
                homePage.VerifyEmailAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC034");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC038_VerifyWhenUserClickOnCancelButtonInOptionsForCreatingYourCalendarReportPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC038-Verify when user click on Cancel button in Options For Creating Your Calendar Report popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Excel");
                calendar.VerifyExportExcelPopup();
                calendar.SelectCheckboxOrClickButtonInExcelPopup("Cancel");
                calendar.VerifyExportExcelPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC038");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC039_40_VerifyWhenUserClickOnEmailAsAttachmentButtonInSendEmailAsPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC039_40-Verify when user click on Email As Attachment button in Send Email As screen.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Email Option");
                calendar.VerifySendEmailAsPopup(true, "Email As Attachment");
                homePage.VerifyAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC039_40");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC041_VerifyWhenUserClickOnEmailWithDownloadLinkButtonInSendEmailAsPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC041-Verify when user click on Email with Download Link button in Send Email As screen.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Email Option");
                calendar.VerifySendEmailAsPopup(true, "Email with Download Link");
                homePage.VerifyAlertPopupMessageAndClickButton("Your requested report(s) are being created, you should receive email(s) shortly. For assistance, contact Numerator at ", "Okay, Got It");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC041");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC042_VerifyWhenUserClickOnCancelButtonInSendEmailAsPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC042-Verify when user click on Cancel button in Send Email As screen.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMenuOptionInCalendarView("Export", "Email Option");
                calendar.VerifySendEmailAsPopup(true, "Cancel");
                calendar.VerifySendEmailAsPopup(false);
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC042");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC043_44_45_VerifyMonthDDL_YearDDLAndCalendarNavigation(String Bname)
        {
            TestFixtureSetUp(Bname, "TC043_44_45-Verify Month DDL, Year DDL and Calendar Navigation.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.SelectMonthFromDDLInCalendar();
                calendar.SelectYearFromDDLInCalendar();
                calendar.VerifyCalendarNavigation(false);
                calendar.VerifyCalendarNavigation();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC043_44_45");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC046_47_VerifyWhenUserClickOnLinkToDrillAdsImage(String Bname)
        {
            TestFixtureSetUp(Bname, "TC046_47-Verify when user click on Link to Drill Ads image.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.VerifyCarouselFromCalendarScreen();
                calendar.VerifyExportAsExcelOptionFromCarousel();
                homePage.VerifyFileDownloadedOrNotOnScreen("Numerator Promotions Intel_Detail_Report_Promoted_Products", "*.xlsx");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC046_47");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC048_VerifyWhenUserClickOnImageNavigationArrowsFromCarousel(String Bname)
        {
            TestFixtureSetUp(Bname, "TC048-Verify when user click on Image Navigation arrows from Carousel.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.VerifyNavigationOnCalendarCarousel();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC048");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC049_VerifyWhenUserClickOnViewAdOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC049-Verify when user click on View Ad option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.VerifyCarouselFromCalendarScreen();
                calendar.OpenViewAdOrDetailPopup(false);
                adDetails.VerifyViewAdPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC049");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC050_VerifyWhenUserClickOnDetailOption(String Bname)
        {
            TestFixtureSetUp(Bname, "TC050-Verify when user click on Detail option.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.VerifyCarouselFromCalendarScreen();
                calendar.OpenViewAdOrDetailPopup(true);
                adDetails.VerifyDetailPopup();
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC050");
                throw;
            }
            driver.Quit();
        }

        [Test]
        [TestCaseSource(typeof(Base), "BrowserToRun")]
        public void TC051_VerifyWhenUserClickOnDownloadTabFromPopup(String Bname)
        {
            TestFixtureSetUp(Bname, "TC051-Verify when user click on Download tab from popup.");
            try
            {
                loginPage.navigateToLoginPage().VerifyLoginPage();
                loginPage.loginUsingValidEmailIdAndPassword(0).clickSignInButton();

                homePage.VerifyHomePage();
                homePage.VerifyLeftNavigationMenuListAndSelectOption("Settings");
                homePage.VerifyClientAndChangeIfItDoesNotMatch("Procter & Gamble");
                homePage.VerifyHomePage();

                homePage.VerifyLeftNavigationMenuListAndSelectOption("Calendar");
                calendar.VerifyCalendarScreen();
                homePage.VerifyHomePage();
                calendar.VerifyCarouselFromCalendarScreen();
                calendar.OpenViewAdOrDetailPopup(false);
                adDetails.VerifyViewAdPopup();
                adDetails.VerifyDownloadFunctionalityInViewAdPopup("Download Current Page");
                homePage.VerifyFileDownloadedOrNotOnScreen("", "*.pdf");
                adDetails.VerifyDownloadFunctionalityInViewAdPopup("Download Selected Page");
                homePage.VerifyFileDownloadedOrNotOnScreen("", "*.pdf");
                adDetails.VerifyDownloadFunctionalityInViewAdPopup("Download Full Ad");
                homePage.VerifyFileDownloadedOrNotOnScreen("", "*.pdf");
            }
            catch (Exception e)
            {
                Logging.LogStop(this.driver, test, e, MethodBase.GetCurrentMethod(), Bname + "_TestSuite007_Calendar_DetailLevel_TC051");
                throw;
            }
            driver.Quit();
        }

        #endregion
    }
}
